using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocxToolbox;

/// <summary>
/// 一次清理请求。图形界面和命令行都填这个结构，再交给 <see cref="WordCleaner.ProcessFile"/>。
/// </summary>
public sealed class CleanRequest
{
    public bool Headers { get; set; }
    public bool LastSection { get; set; }
    public bool Props { get; set; }
    public bool Privacy { get; set; }
    public bool Orphans { get; set; }
    public string Algo { get; set; } = "MD5";

    /// <summary>要删除的图片哈希（小写十六进制），可以给多个。</summary>
    public List<string> ImageHashes { get; } = new();

    /// <summary>删除最后一节时忽略"空白"判断，连内容一起删（对应 --force）。</summary>
    public bool ForceLastSection { get; set; }

    /// <summary>接受所有修订并清除批注（会改变文档可见内容，对应 --track-changes）。</summary>
    public bool TrackChanges { get; set; }

    /// <summary>
    /// 水印/来源检测的触发词（对应 --watermark）。非空即启用：命中后只清理「水印对象」，
    /// 不碰正文插图。触发词完全由用户提供，程序不预置任何品牌词。
    /// </summary>
    public List<string> WatermarkKeywords { get; } = new();

    /// <summary>命中后连正文插图一起删（对应 --wm-drop-all-images）。激进，慎用。</summary>
    public bool WatermarkDropAllImages { get; set; }

    /// <summary>本机 ppocr 目录，给了才会对图片做 OCR（对应 --ocr）。</summary>
    public string? OcrDir { get; set; }

    /// <summary>每个文档最多 OCR 多少张图，0 表示不限（对应 --ocr-max）。</summary>
    public int OcrMaxImages { get; set; } = 60;

    /// <summary>是否至少指定了一个清理项。</summary>
    public bool HasWork => Headers || LastSection || Props || Privacy || Orphans || ImageHashes.Count > 0
                           || TrackChanges || WatermarkKeywords.Count > 0;
}

/// <summary>删除最后一节的结果。</summary>
public enum LastSectionResult
{
    /// <summary>删掉了。</summary>
    Removed,
    /// <summary>文档只有一节，删了等于删全文。</summary>
    SingleSection,
    /// <summary>最后一节还有内容，按"只删空白节"的约定不动它。</summary>
    NotBlank,
}

/// <summary>
/// 水印/来源检测的结果。Hit 为 true 表示文档里确实出现了指定文字（正文文字、VML 水印属性或图片 OCR）。
/// </summary>
public sealed class WatermarkReport
{
    /// <summary>是否命中任何触发词。</summary>
    public bool Hit { get; set; }

    /// <summary>是否是「文字/属性」命中（区别于只有图片 OCR 命中）。</summary>
    public bool TextHit { get; set; }

    /// <summary>是否执行了「删掉全部图片」的激进策略。</summary>
    public bool DroppedAllImages { get; set; }

    /// <summary>删掉的水印图形元素个数（VML shape / DrawingML drawing）。</summary>
    public int ShapesRemoved { get; set; }

    /// <summary>删掉的关键词文字个数（正文、页眉页脚、脚注尾注里的命中文字，只删关键词这几个字）。</summary>
    public int KeywordsRemoved { get; set; }

    /// <summary>顺带清掉的媒体部件个数（不再被引用）。</summary>
    public int ImagesRemoved { get; set; }

    /// <summary>实际送去做 OCR 的图片张数。</summary>
    public int OcrScanned { get; set; }

    /// <summary>文档里的图片元素总数（DrawingML blip + VML imagedata），只在只读探测时填充。</summary>
    public int ImageElements { get; set; }

    /// <summary>OCR 没能跑起来的原因；为空表示没请求或跑成功。</summary>
    public string? OcrNote { get; set; }

    private readonly List<string> _evidence = new();

    /// <summary>命中的证据（截断后的片段），最多留 6 条，供日志展示。</summary>
    public IReadOnlyList<string> Evidence => _evidence;

    public void AddEvidence(string text)
    {
        if (_evidence.Count >= 6 || _evidence.Contains(text)) return;
        _evidence.Add(text);
    }
}

/// <summary>一个节的分析结果（只读，供探测和删除共用）。</summary>
public sealed class DocumentSection
{
    public int Number { get; init; }
    /// <summary>该节第一个块在 body 顶层块里的序号（从 1 开始，含）。</summary>
    public int FirstBlock { get; init; }
    /// <summary>该节最后一个块的序号（含）。</summary>
    public int LastBlock { get; init; }
    public int Paragraphs { get; init; }
    public int Tables { get; init; }
    public int HeaderRefs { get; init; }
    public int FooterRefs { get; init; }
    /// <summary>节里有没有实质内容（文字、表格、图片、批注等）。</summary>
    public bool HasContent { get; init; }
    public bool IsLast { get; init; }
    /// <summary>承载本节 sectPr 的顶层块；末尾节的 sectPr 挂在 body 上，所以是 null。</summary>
    public OpenXmlElement? BoundaryBlock { get; init; }
    public SectionProperties? SectPr { get; init; }
}

/// <summary>
/// Word 文档清理的核心逻辑，供图形界面（MainForm）和命令行（Cli）共用。
/// </summary>
public static class WordCleaner
{
    /// <summary>处理后写盘用的临时文件后缀（同目录、同卷，才能用 File.Replace 原子替换）。</summary>
    private const string TempSuffix = ".wtmp";

    /// <summary>关系命名空间：XML 里所有 r:id / r:embed / r:link 都在这个 ns 下。</summary>
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>WordprocessingML 主命名空间（属性剥离时只认这个命名空间，避免误伤其他格式）。</summary>
    private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// 修订/批注标记里「连元素一起删」的本地名。删 w:del 时其 delText 内容随元素消失
    /// （等价于接受修订）；range 标记和格式修订（*Change）都是空壳，直接删。
    /// 批注锚点也走这条路，只是单独计数。
    /// </summary>
    private static readonly HashSet<string> DeletedRevisionElements = new(StringComparer.Ordinal)
    {
        // 文字/单元格修订：del 的内容在内部，随元素删除；cellDel/cellMerge 是空标记
        "del", "cellDel", "cellMerge",
        // 移动修订：moveFrom 的内容随元素删除，RangeStart/End 是空标记
        "moveFrom", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd",
        // 不一致修订的范围标记（ins/del 成段时配套出现）
        "insRangeStart", "insRangeEnd", "delRangeStart", "delRangeEnd",
        // 格式修订：改过哪些格式没有保留价值，删
        "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange",
        "tcPrChange", "tblGridChange", "numberingChange",
        // customXml 里的修订标记
        "customXmlIns", "customXmlDel", "customXmlMoveFrom", "customXmlMoveTo",
        "customXmlInsRangeStart", "customXmlInsRangeEnd",
        "customXmlDelRangeStart", "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
        // 批注锚点：同批必须删掉 comments 部件（见 RemoveTrackChangesAndComments）
        "commentRangeStart", "commentRangeEnd", "commentReference",
    };

    /// <summary>修订标记里「解包保留内容」的本地名：把子元素原位置提升，只去掉包裹层。</summary>
    private static readonly HashSet<string> UnwrappedRevisionElements = new(StringComparer.Ordinal)
    {
        "ins", "moveTo", "cellIns",
    };

    /// <summary>修订/批注相关的身份属性（Word 命名空间），全树剥离。</summary>
    private static readonly HashSet<string> StrippedRevisionAttributes = new(StringComparer.Ordinal)
    {
        "author", "date", "userId", "initials", "ed",
    };

    /// <summary>支持的图片哈希算法。</summary>
    public static readonly string[] SupportedAlgorithms = { "MD5", "SHA1", "SHA256" };

    /// <summary>把算法名规范成受支持的值，未知值回退为 MD5。</summary>
    public static string NormalizeAlgorithm(string? algo) => algo switch
    {
        "SHA1"   => "SHA1",
        "SHA256" => "SHA256",
        _        => "MD5",
    };

    /// <summary>
    /// 处理一个文件：先复制到同目录临时文件，改动并校验通过后再原子替换原文件。
    /// 任何一步失败（打开失败、校验错误增加、文件被占用）都保留原文件不变。
    /// 返回值只表示"没有出错"，被跳过（没有需要改的内容）也算成功。
    /// </summary>
    public static bool ProcessFile(string filePath, CleanRequest req, Action<string> log)
    {
        string fullPath = Path.GetFullPath(filePath);
        string dir = Path.GetDirectoryName(fullPath) ?? ".";
        string temp = Path.Combine(dir, "." + Path.GetFileName(fullPath) + TempSuffix);

        try
        {
            // 预检 1：只读文件写回去必然失败，提前报出来，别做完再失败
            if ((File.GetAttributes(fullPath) & FileAttributes.ReadOnly) != 0)
            {
                log($"[异常] 文件是只读的，已跳过: {fullPath}");
                return false;
            }

            // 预检 2：被 Word 打开的文件（共享读占用）替换会失败
            if (!CanOpenExclusive(fullPath, out string? lockReason))
            {
                log($"[异常] 文件被占用或不可写，已跳过: {fullPath} - {lockReason}");
                return false;
            }

            // 内容先读进内存：校验用内存副本，损坏的包不会锁住磁盘上的临时文件
            byte[] data = File.ReadAllBytes(fullPath);

            // 校验基线：原文档本身可能已有 OpenXML 错误，比对的是"有没有变差"
            int baseline = CountValidationErrors(data, out string? reason);
            if (baseline < 0)
            {
                log($"[异常] 无法作为 .docx 打开（{reason}），已保留原文件: {fullPath}");
                return false;
            }

            File.WriteAllBytes(temp, data);

            bool changed = false;

            using (var doc = WordprocessingDocument.Open(temp, true))
            {
                if (req.Headers)
                {
                    int parts = RemoveHeadersAndFooters(doc);
                    if (parts > 0)
                    {
                        log($"[成功] 删除页眉页脚: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 没有页眉页脚: {fullPath}");
                    }
                }

                if (req.LastSection)
                {
                    var (removed, stop) = RemoveLastSection(doc, req.ForceLastSection);
                    if (removed > 0)
                    {
                        string what = req.ForceLastSection
                            ? "强制删除最后一节（含内容）"
                            : (removed == 1 ? "删除末尾空白节" : $"删除末尾空白节 {removed} 个");
                        log($"[成功] {what}: {fullPath}");
                        changed = true;
                    }
                    else if (stop == LastSectionResult.SingleSection)
                    {
                        log($"[跳过] 文档仅有一节，无法删除: {fullPath}");
                    }
                    else
                    {
                        log($"[跳过] 最后一节还有内容，按「只删空白节」的约定未删除: {fullPath}");
                    }
                }

                if (req.Props)
                {
                    if (RemovePropertiesAndInfo(doc))
                    {
                        log($"[成功] 删除属性信息: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 属性信息已经为空: {fullPath}");
                    }
                }

                if (req.Privacy)
                {
                    int marks = RemovePrivacyResidue(doc, m => log("    " + m));
                    if (marks > 0)
                    {
                        log($"[成功] 清除隐私残留标记 {marks} 处: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 没有可清除的隐私标记: {fullPath}");
                    }
                }

                if (req.TrackChanges)
                {
                    int marks = RemoveTrackChangesAndComments(doc, m => log("    " + m));
                    if (marks > 0)
                    {
                        log($"[成功] 接受修订、清除批注 {marks} 处: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 没有可清除的修订或批注: {fullPath}");
                    }
                }

                // 水印/来源检测放在「按哈希删图」之前：先按关键词清掉水印对象，再处理用户点名的图片
                if (req.WatermarkKeywords.Count > 0)
                {
                    var wm = RemoveWatermarks(doc, req.WatermarkKeywords, req.WatermarkDropAllImages,
                                              req.OcrDir, req.OcrMaxImages, m => log("    " + m));

                    if (!wm.Hit)
                    {
                        log($"[跳过] 未检出指定文字，图片未动: {fullPath}");
                    }
                    else
                    {
                        string why = wm.Evidence.Count > 0
                            ? "（命中：" + string.Join("；", wm.Evidence.Take(3))
                              + (wm.Evidence.Count > 3 ? $"…共 {wm.Evidence.Count} 类" : "") + "）"
                            : "";
                        if (wm.DroppedAllImages)
                        {
                            log($"[成功] 命中关键词，删除全部图片 {wm.ShapesRemoved} 处、命中文字 {wm.KeywordsRemoved} 处、媒体 {wm.ImagesRemoved} 个{why}: {fullPath}");
                            changed = true;
                        }
                        else if (wm.ShapesRemoved > 0 || wm.ImagesRemoved > 0 || wm.KeywordsRemoved > 0)
                        {
                            log($"[成功] 命中关键词，删除水印对象 {wm.ShapesRemoved} 处、命中文字 {wm.KeywordsRemoved} 处、媒体 {wm.ImagesRemoved} 个{why}: {fullPath}");
                            changed = true;
                        }
                        else
                        {
                            log($"[跳过] 命中关键词，但没找到可删的水印对象{why}: {fullPath}");
                        }
                    }
                }

                if (req.ImageHashes.Count > 0)
                {
                    int count = RemoveImagesByHashes(doc, req.ImageHashes, req.Algo, m => log("    " + m));
                    if (count > 0)
                    {
                        log($"[成功] 删除匹配图片 {count} 处: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 未找到哈希匹配的图片: {fullPath}");
                    }
                }

                // 放在最后：前面各步删掉引用后留下的媒体在这里一次性收走
                if (req.Orphans)
                {
                    var (files, bytes) = RemoveOrphanMedia(doc);
                    if (files > 0)
                    {
                        log($"[成功] 清理未引用的媒体 {files} 个，节省 {FormatSize(bytes)}: {fullPath}");
                        changed = true;
                    }
                    else
                    {
                        log($"[跳过] 没有未引用的媒体: {fullPath}");
                    }
                }
            }

            // 包级清扫：删掉部件后常留下不可达的 zip 条目（如 customXml 的 itemProps），
            // SDK 不管这类部件，只能在包关闭后从 zip 层面判断可达性
            if (req.Orphans || req.Privacy || req.TrackChanges)
            {
                var (removedEntries, freedBytes) = RemoveUnreachablePackageEntries(temp, m => log("    " + m));
                if (removedEntries > 0)
                {
                    log($"[成功] 清理包内失效条目 {removedEntries} 个，节省 {FormatSize(freedBytes)}: {fullPath}");
                    changed = true;
                }
            }

            if (!changed)
            {
                log($"[跳过] 没有需要修改的内容: {fullPath}");
                TryDelete(temp);
                return true;
            }

            int after = CountValidationErrors(File.ReadAllBytes(temp), out string? afterReason);
            if (after < 0)
            {
                log($"[异常] 处理后无法校验（{afterReason}），已放弃写入并保留原文件: {fullPath}");
                TryDelete(temp);
                return false;
            }
            if (after > baseline)
            {
                log($"[异常] 处理后校验错误由 {baseline} 增至 {after}，已放弃写入并保留原文件: {fullPath}");
                TryDelete(temp);
                return false;
            }

            try
            {
                File.Replace(temp, fullPath, null);
            }
            catch (Exception ex)
            {
                TryDelete(temp);
                log($"[异常] 写入原文件失败（文件内容未改变）: {fullPath} - {ex.Message}");
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            TryDelete(temp);
            log($"[异常] 处理失败: {fullPath} - {ex.Message}");
            return false;
        }
    }

    /// <summary>探测文件能否以独占方式打开——用于提前发现被 Word 占用/无写权限的文件。</summary>
    private static bool CanOpenExclusive(string path, out string? reason)
    {
        reason = null;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return true;
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            return false;
        }
    }

    /// <summary>只读：统计一个 .docx 的 OpenXML 校验错误数；无法打开返回 -1，reason 给出原因。</summary>
    public static int ValidateDocument(string filePath, out string? reason)
    {
        try
        {
            return CountValidationErrors(File.ReadAllBytes(filePath), out reason);
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            return -1;
        }
    }

    /// <summary>统计一个文档的 OpenXML 校验错误数；无法打开或校验抛异常时返回 -1。</summary>
    private static int CountValidationErrors(byte[] data, out string? reason)
    {
        reason = null;
        try
        {
            // 用内存流而不是文件：损坏的包在打开失败时会漏出文件句柄，那会让临时文件删不掉
            using var ms = new MemoryStream(data, writable: false);
            using var doc = WordprocessingDocument.Open(ms, false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            return validator.Validate(doc).Count();
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            return -1;
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch
        {
            // 临时文件删不掉不影响结果，下次处理会被覆盖
        }
    }

    /// <summary>删除文档中的全部页眉页脚部件，并清除所有分节里的引用。返回删掉的部件数。</summary>
    public static int RemoveHeadersAndFooters(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart;
        if (mainPart == null) return 0;

        int parts = 0;

        foreach (var headerPart in mainPart.HeaderParts.ToList())
        {
            mainPart.DeletePart(headerPart);
            parts++;
        }

        foreach (var footerPart in mainPart.FooterParts.ToList())
        {
            mainPart.DeletePart(footerPart);
            parts++;
        }

        var body = mainPart.Document.Body;
        if (body == null) return parts;

        // 用 Descendants 而不是 body 的直接子元素：分节属性也可能出现在表格单元格段落里
        foreach (var sectPr in body.Descendants<SectionProperties>().ToList())
        {
            foreach (var r in sectPr.Elements<HeaderReference>().ToList()) r.Remove();
            foreach (var r in sectPr.Elements<FooterReference>().ToList()) r.Remove();
        }

        return parts;
    }

    /// <summary>
    /// 从末尾往前删空白节，一直删到"再删就会伤到有内容的节"或"只剩一节"为止。
    /// 空白指没有任何实质内容（文字、表格、图片、批注、域都算内容）。
    /// force=true 时退回旧行为：只强制删掉最后一节，不管里面有什么。
    /// 每删一节都会把前一节的 sectPr 提升到 body 末尾，页面设置跟着正文走、不会丢。
    /// 返回删掉的节数和停止原因（Removed 表示有删除，其余值在 Removed=0 时用来区分"为什么没删"）。
    /// </summary>
    public static (int Removed, LastSectionResult Stop) RemoveLastSection(WordprocessingDocument doc, bool force = false)
    {
        int removed = 0;

        while (true)
        {
            var body = doc.MainDocumentPart?.Document.Body;
            if (body == null) return (removed, LastSectionResult.SingleSection);

            var sections = ScanSections(doc);
            if (sections.Count <= 1) return (removed, LastSectionResult.SingleSection);

            var last = sections[^1];
            if (!force && last.HasContent) return (removed, LastSectionResult.NotBlank);

            RemoveOneSection(body, sections);
            removed++;

            // --force 保持旧语义：只处理最末尾那一节，剩下几节是什么内容留给用户自己判断
            if (force) return (removed, LastSectionResult.Removed);
        }
    }

    /// <summary>删掉 sections 里最后一节的块，并把前一节的 sectPr 挂到 body 末尾。</summary>
    private static void RemoveOneSection(Body body, List<DocumentSection> sections)
    {
        var last = sections[^1];
        var blocks = body.ChildElements.Where(e => e is not SectionProperties).ToList();

        // 先把前一节的 sectPr 从它所在的段落里摘出来，待会儿挂到 body 末尾
        var prevSectPr = sections[^2].SectPr;
        prevSectPr?.Remove();

        // 删掉末尾节的所有块
        foreach (var block in blocks.Skip(last.FirstBlock - 1).ToList())
        {
            block.Remove();
        }

        // 删掉 body 上原有的末尾 sectPr，换上前一节的
        body.Elements<SectionProperties>().LastOrDefault()?.Remove();
        if (prevSectPr != null)
        {
            body.AppendChild(prevSectPr);
        }
    }

    /// <summary>
    /// 按文档顺序切分节。段落里的 sectPr 表示"本节到这段为止"，body 末尾的 sectPr 属于最后一节。
    /// 分节符藏在表格、内容控件里也能找到——探测和删除都用这一份结果，避免两边口径不一致。
    /// </summary>
    public static List<DocumentSection> ScanSections(WordprocessingDocument doc)
    {
        var result = new List<DocumentSection>();
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return result;

        var blocks = body.ChildElements.Where(e => e is not SectionProperties).ToList();
        var finalSectPr = body.Elements<SectionProperties>().LastOrDefault();

        // 每个顶层块里藏着哪一段落级 sectPr（没有就是 null）
        var blockSectPr = blocks.Select(b => b.Descendants<SectionProperties>().LastOrDefault()).ToList();

        int start = 0;
        for (int i = 0; i < blocks.Count; i++)
        {
            if (blockSectPr[i] == null) continue;

            result.Add(BuildSection(result.Count + 1, blocks, start, i, blockSectPr[i], false));
            start = i + 1;
        }

        if (start < blocks.Count)
        {
            result.Add(BuildSection(result.Count + 1, blocks, start, blocks.Count - 1, finalSectPr, true));
        }
        else if (result.Count > 0)
        {
            // 退化情况：最后一段自带 sectPr、body 上没有独立 sectPr
            var last = result[^1];
            result[^1] = BuildSection(last.Number, blocks, last.FirstBlock - 1, last.LastBlock - 1,
                                      last.SectPr ?? finalSectPr, true);
        }

        return result;
    }

    private static DocumentSection BuildSection(int number, List<OpenXmlElement> blocks,
                                               int start, int end, SectionProperties? sectPr, bool isLast)
    {
        var slice = blocks.Skip(start).Take(end - start + 1).ToList();

        return new DocumentSection
        {
            Number = number,
            FirstBlock = start + 1,
            LastBlock = end + 1,
            Paragraphs = slice.Count(b => b is Paragraph),
            Tables = slice.Count(b => b is Table),
            HeaderRefs = sectPr?.Elements<HeaderReference>().Count() ?? 0,
            FooterRefs = sectPr?.Elements<FooterReference>().Count() ?? 0,
            HasContent = slice.Any(HasRealContent),
            IsLast = isLast,
            BoundaryBlock = isLast ? null : blocks[end],
            SectPr = sectPr,
        };
    }

    /// <summary>块里有没有实质内容：纯空段落不算，表格、图片、批注、脚注、域都算。</summary>
    private static bool HasRealContent(OpenXmlElement block)
    {
        if (block is Table) return true;
        if (block.Descendants<Table>().Any()) return true;

        if (block.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text))) return true;
        if (block.Descendants<Drawing>().Any()) return true;

        foreach (var d in block.Descendants())
        {
            switch (d.LocalName)
            {
                case "pict":
                case "object":
                case "oleObject":
                case "AlternateContent":
                    return true;
            }
        }

        if (block.Descendants<CommentReference>().Any()) return true;
        if (block.Descendants<FootnoteReference>().Any()) return true;
        if (block.Descendants<EndnoteReference>().Any()) return true;
        if (block.Descendants<SimpleField>().Any()) return true;
        if (block.Descendants<FieldCode>().Any()) return true;

        return false;
    }

    /// <summary>
    /// 清除可识别身份的属性：核心属性（作者、标题、时间戳等）、扩展属性部件、
    /// 自定义属性部件（docProps/custom.xml）和缩略图（docProps/thumbnail.jpeg）。
    /// 返回是否真的改动了什么。
    /// </summary>
    public static bool RemovePropertiesAndInfo(WordprocessingDocument doc)
    {
        bool changed = false;
        var props = doc.PackageProperties;

        if (!string.IsNullOrEmpty(props.Creator))        { props.Creator = string.Empty;        changed = true; }
        if (!string.IsNullOrEmpty(props.LastModifiedBy)) { props.LastModifiedBy = string.Empty; changed = true; }
        if (!string.IsNullOrEmpty(props.Title))          { props.Title = string.Empty;          changed = true; }
        if (!string.IsNullOrEmpty(props.Subject))        { props.Subject = string.Empty;        changed = true; }
        if (!string.IsNullOrEmpty(props.Keywords))       { props.Keywords = string.Empty;       changed = true; }
        if (!string.IsNullOrEmpty(props.Description))    { props.Description = string.Empty;    changed = true; }
        if (!string.IsNullOrEmpty(props.Category))       { props.Category = string.Empty;       changed = true; }
        if (!string.IsNullOrEmpty(props.ContentStatus))  { props.ContentStatus = string.Empty;  changed = true; }
        if (!string.IsNullOrEmpty(props.Version))        { props.Version = string.Empty;        changed = true; }
        if (!string.IsNullOrEmpty(props.Revision))       { props.Revision = string.Empty;       changed = true; }

        if (props.Created != null)    { props.Created = null;    changed = true; }
        if (props.Modified != null)   { props.Modified = null;   changed = true; }
        if (props.LastPrinted != null){ props.LastPrinted = null; changed = true; }

        if (doc.ExtendedFilePropertiesPart != null)
        {
            doc.DeletePart(doc.ExtendedFilePropertiesPart);
            changed = true;
        }

        if (doc.CustomFilePropertiesPart != null)
        {
            doc.DeletePart(doc.CustomFilePropertiesPart);
            changed = true;
        }

        if (doc.ThumbnailPart != null)
        {
            doc.DeletePart(doc.ThumbnailPart);
            changed = true;
        }

        return changed;
    }

    /// <summary>
    /// 清除能反推来源的残余标记，返回清除项数量：
    /// rsid 编辑会话 ID、settings 里的 rsids / docVars / attachedTemplate、
    /// customXml 存储（连同内容控件的数据绑定）、外部超链接（保留文字）。
    /// </summary>
    public static int RemovePrivacyResidue(WordprocessingDocument doc, Action<string>? log = null)
    {
        var main = doc.MainDocumentPart;
        if (main == null) return 0;

        int changes = 0;
        var parts = CollectContentParts(doc);

        // 1) rsid：段落、表格、分节上的编辑会话编号，能反推谁在什么时候改过
        foreach (var part in parts)
        {
            var root = part.RootElement;
            if (root == null) continue;
            int n = StripRsidAttributes(root);
            if (n > 0) log?.Invoke($"清除 {part.GetType().Name} 中的 rsid 属性 {n} 个");
            changes += n;
        }

        // 2) 设置里的模板路径、文档变量、整块编辑会话记录
        var settingsPart = main.DocumentSettingsPart;
        var settings = settingsPart?.Settings;
        if (settingsPart != null && settings != null)
        {
            foreach (var el in settings.ChildElements.Where(e => e.LocalName == "docVars").ToList())
            {
                el.Remove();
                changes++;
                log?.Invoke("删除文档变量 docVars");
            }

            foreach (var el in settings.ChildElements.Where(e => e.LocalName == "rsids").ToList())
            {
                el.Remove();
                changes++;
                log?.Invoke("删除编辑会话记录 rsids");
            }

            foreach (var el in settings.ChildElements.Where(e => e.LocalName == "attachedTemplate").ToList())
            {
                string? rid = el.GetAttributes()
                                .Where(a => a.LocalName == "id" && a.NamespaceUri == RelNs)
                                .Select(a => a.Value)
                                .FirstOrDefault();
                el.Remove();
                changes++;
                if (!string.IsNullOrEmpty(rid)) DeleteRelationship(settingsPart, rid);
                log?.Invoke("删除附加模板引用 attachedTemplate");
            }
        }

        // 3) customXml：模板/插件常把路径、工号、导出数据塞在这里
        var customXmlParts = main.CustomXmlParts.ToList();
        if (customXmlParts.Count > 0)
        {
            // 先摘数据绑定，避免内容控件指向已删除的存储
            foreach (var part in parts)
            {
                var root = part.RootElement;
                if (root == null) continue;
                foreach (var binding in root.Descendants().Where(d => d.LocalName == "dataBinding").ToList())
                {
                    binding.Remove();
                    changes++;
                }
            }

            foreach (var cx in customXmlParts)
            {
                // 先删子部件（itemProps、图片等）：DeletePart 只管自己那一层关系，
                // 先删父部件会把子部件留在包里变成真孤儿
                foreach (var child in cx.Parts.ToList())
                {
                    try { cx.DeletePart(child.OpenXmlPart); } catch { /* 子部件删不掉不阻塞 */ }
                }

                main.DeletePart(cx);
                changes++;
            }
            log?.Invoke($"删除 customXml 部件 {customXmlParts.Count} 个");
        }

        // 4) 外部超链接：保留文字和格式，只摘掉指向（常带内部共享盘路径）
        foreach (var part in parts)
        {
            var links = part.HyperlinkRelationships.ToList();
            if (links.Count == 0) continue;

            var root = part.RootElement;
            var ids = new HashSet<string>(links.Select(l => l.Id), StringComparer.Ordinal);

            if (root != null)
            {
                foreach (var h in root.Descendants<Hyperlink>().ToList())
                {
                    var id = h.Id?.Value;
                    if (id != null && ids.Contains(id))
                    {
                        h.Id = null;
                        changes++;
                    }
                }
            }

            foreach (var link in links)
            {
                DeleteRelationship(part, link.Id);
                changes++;
            }
            log?.Invoke($"摘除外部超链接 {links.Count} 个（{part.GetType().Name}）");
        }

        return changes;
    }

    /// <summary>
    /// 接受所有修订并清除批注，返回改动项总数。
    /// 修订标记可能出现在页眉页脚、脚注、尾注里，扫描范围不能只看 document.xml；
    /// 批注锚点与 comments 部件必须同批处理——只删锚点不删部件，Word 会报文档损坏。
    /// 语义等价于「接受所有修订」：w:ins/w:moveTo 的内容保留（解包），w:del/w:moveFrom
    /// 的内容真的消失。因此这是一个独立选项，不并入 --privacy 或 --all。
    /// </summary>
    public static int RemoveTrackChangesAndComments(WordprocessingDocument doc, Action<string>? log = null)
    {
        var main = doc.MainDocumentPart;
        if (main == null) return 0;

        int removed = 0;
        int unwrapped = 0;
        int comments = 0;
        int attrStripped = 0;

        // 1) 正文、页眉、页脚、脚注、尾注：处理修订标记与批注锚点
        foreach (var part in CollectContentParts(doc))
        {
            var root = part.RootElement;
            if (root == null) continue;

            // 先拍快照再逐个处理：树在动，foreach 原集合会漏项
            foreach (var el in root.Descendants().ToList())
            {
                if (el.Parent == null) continue; // 已被随祖先一起删掉（如 w:del 里的 w:ins）

                string name = el.LocalName;

                if (DeletedRevisionElements.Contains(name))
                {
                    if (name.StartsWith("comment", StringComparison.Ordinal)) comments++;
                    else removed++;
                    el.Remove();
                    continue;
                }

                if (UnwrappedRevisionElements.Contains(name))
                {
                    UnwrapElement(el);
                    unwrapped++;
                }
            }

            attrStripped += StripRevisionAttributes(root);
        }

        if (removed > 0)  log?.Invoke($"删除修订标记 {removed} 处（del/moveFrom/格式修订等）");
        if (unwrapped > 0) log?.Invoke($"解包保留内容 {unwrapped} 处（ins/moveTo/cellIns）");
        if (comments > 0) log?.Invoke($"清除批注锚点 {comments} 处");
        if (attrStripped > 0) log?.Invoke($"剥离修订身份属性 {attrStripped} 个（author/date/userId 等）");

        // 2) 批注部件必须整删：锚点清掉后，comments.xml（及 commentsExtended/commentsIds/people）
        //    及其关系一起移除，否则 Word 打开会报「文档损坏」。
        int partsDeleted = DeleteCommentsLikeParts(main, log);
        comments += partsDeleted;

        return removed + unwrapped + comments + attrStripped;
    }

    /// <summary>把元素解包：子元素按原顺序提升到父节点，去掉包裹层（如 w:ins 的外壳）。</summary>
    private static void UnwrapElement(OpenXmlElement element)
    {
        if (element.Parent == null) return;

        foreach (var child in element.ChildElements.ToList())
        {
            element.RemoveChild(child);
            element.InsertBeforeSelf(child);
        }
        element.Remove();
    }

    /// <summary>全树剥离修订相关的身份属性（author/date/userId/initials/ed，仅 Word 命名空间）。</summary>
    private static int StripRevisionAttributes(OpenXmlElement root)
    {
        int n = 0;
        foreach (var el in root.Descendants().Prepend(root).ToList())
        {
            foreach (var attr in el.GetAttributes()
                                   .Where(a => a.NamespaceUri == WordNs
                                            && StrippedRevisionAttributes.Contains(a.LocalName))
                                   .ToList())
            {
                el.RemoveAttribute(attr.LocalName, attr.NamespaceUri);
                n++;
            }
        }
        return n;
    }

    /// <summary>整删批注/人员相关部件（comments/commentsExtended/commentsIds/people），返回删掉的部件数。</summary>
    private static int DeleteCommentsLikeParts(OpenXmlPartContainer root, Action<string>? log)
    {
        // 按 content type 认部件（跨 SDK 版本最稳，不依赖各部件类型的强类型属性）
        bool IsCommentsLike(OpenXmlPart part)
        {
            string ct = part.ContentType ?? "";
            return ct.Contains("comments", StringComparison.OrdinalIgnoreCase)
                || ct.Contains("people", StringComparison.OrdinalIgnoreCase);
        }

        // 收集所有后代部件及其父容器（DeletePart 必须从持有关系的容器上调用）
        var doomed = new List<(OpenXmlPartContainer Parent, OpenXmlPart Part)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var queue = new Queue<(OpenXmlPartContainer Parent, OpenXmlPart Part)>();
        foreach (var pair in root.Parts) queue.Enqueue((root, pair.OpenXmlPart));

        while (queue.Count > 0)
        {
            var (parent, part) = queue.Dequeue();
            if (!seen.Add(part.Uri.OriginalString)) continue;
            if (IsCommentsLike(part)) doomed.Add((parent, part));
            foreach (var pair in part.Parts) queue.Enqueue((part, pair.OpenXmlPart));
        }

        // 先删最深的（commentsIds 挂在 comments 下面），再删 comments 本体。
        // 注意：DeletePart 之后 part 就被销毁了，日志里不能再访问 part.Uri，必须事先取出字符串。
        int deleted = 0;
        foreach (var (parent, part) in Enumerable.Reverse(doomed))
        {
            string uri = part.Uri.OriginalString;
            try
            {
                parent.DeletePart(part);
                deleted++;
                log?.Invoke($"删除批注相关部件: {uri}");
            }
            catch (Exception ex)
            {
                log?.Invoke($"[警告] 删除部件失败: {uri} - {ex.Message}");
            }
        }

        return deleted;
    }

    /// <summary>
    /// 删除"有关系但没有任何 XML 元素引用"的媒体部件（图片被删、页眉被删之后常见的残留）。
    /// 返回（删除个数，节省字节数）。
    /// </summary>
    public static (int Files, long Bytes) RemoveOrphanMedia(WordprocessingDocument doc)
    {
        int removed = 0;
        long bytes = 0;

        var main = doc.MainDocumentPart;
        if (main == null) return (0, 0);

        foreach (var part in CollectContentParts(doc))
        {
            var root = part.RootElement;
            if (root == null) continue;

            // 这个部件里出现过的所有关系 ID（r:id / r:embed / r:link 等）
            var used = new HashSet<string>(StringComparer.Ordinal);
            foreach (var el in root.Descendants().Prepend(root))
            {
                foreach (var attr in el.GetAttributes())
                {
                    if (attr.NamespaceUri == RelNs && !string.IsNullOrEmpty(attr.Value))
                    {
                        used.Add(attr.Value);
                    }
                }
            }

            foreach (var pair in part.Parts.ToList())
            {
                if (pair.OpenXmlPart is not ImagePart image) continue;
                if (used.Contains(pair.RelationshipId)) continue;

                long size = 0;
                try
                {
                    using var s = image.GetStream(FileMode.Open, FileAccess.Read);
                    size = s.Length;
                }
                catch
                {
                    // 读不到长度不影响删除
                }

                try
                {
                    part.DeletePart(image);
                    removed++;
                    bytes += size;
                }
                catch
                {
                    // 删不掉就算了，不影响文档本身
                }
            }
        }

        return (removed, bytes);
    }

    /// <summary>正文、页眉、页脚、脚注、尾注这几个可能带图片和标记的部件。</summary>
    private static List<OpenXmlPart> CollectContentParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        var parts = new List<OpenXmlPart>();
        if (main == null) return parts;

        parts.Add(main);
        parts.AddRange(main.HeaderParts);
        parts.AddRange(main.FooterParts);
        if (main.FootnotesPart != null) parts.Add(main.FootnotesPart);
        if (main.EndnotesPart != null) parts.Add(main.EndnotesPart);
        return parts;
    }

    private static int StripRsidAttributes(OpenXmlElement root)
    {
        int n = 0;
        foreach (var el in root.Descendants().Prepend(root).ToList())
        {
            var attrs = el.GetAttributes()
                          .Where(a => a.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase))
                          .ToList();

            foreach (var attr in attrs)
            {
                el.RemoveAttribute(attr.LocalName, attr.NamespaceUri);
                n++;
            }
        }
        return n;
    }

    private static void DeleteRelationship(OpenXmlPartContainer container, string id)
    {
        try { container.DeleteExternalRelationship(id); } catch { /* 不是外部关系 */ }
        try { container.DeleteReferenceRelationship(id); } catch { /* 不是内部关系 */ }
    }

    private static string FormatSize(long bytes) =>
        bytes >= 1024 * 1024 ? $"{bytes / 1024.0 / 1024.0:0.0} MB"
        : bytes >= 1024 ? $"{bytes / 1024.0:0.0} KB"
        : $"{bytes} B";

    /// <summary>
    /// 包级清扫：从 _rels/.rels 出发按关系图做可达性分析，删掉没人引用的 zip 条目
    /// （部件删除后遗留的 itemProps、孤儿 rels 等）。顺带剪掉 [Content_Types].xml 里
    /// 指向已不存在部件的 Override。返回（删除条目数，节省字节数）。
    /// </summary>
    private static (int Removed, long Bytes) RemoveUnreachablePackageEntries(string path, Action<string>? log)
    {
        int removed = 0;
        long bytes = 0;

        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var names = new HashSet<string>(zip.Entries.Select(e => e.FullName), StringComparer.OrdinalIgnoreCase);

        var reachable = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var relsKept = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "_rels/.rels" };

        var queue = new Queue<string>(ReadRelationshipTargets(zip, "_rels/.rels", ""));
        while (queue.Count > 0)
        {
            string part = queue.Dequeue();
            if (!names.Contains(part) || !reachable.Add(part)) continue;

            string relsPath = RelationshipPartPath(part);
            if (!names.Contains(relsPath)) continue;
            relsKept.Add(relsPath);

            foreach (string target in ReadRelationshipTargets(zip, relsPath, PartDirectory(part)))
            {
                queue.Enqueue(target);
            }
        }

        var keep = new HashSet<string>(reachable, StringComparer.OrdinalIgnoreCase);
        keep.UnionWith(relsKept);
        keep.Add("[Content_Types].xml");

        foreach (var entry in zip.Entries.ToList())
        {
            if (keep.Contains(entry.FullName)) continue;

            long len = entry.Length;
            string name = entry.FullName;
            entry.Delete();
            removed++;
            bytes += len;
            log?.Invoke($"移除失效条目: {name}");
        }

        PruneContentTypeOverrides(zip, keep);

        return (removed, bytes);
    }

    /// <summary>把 [Content_Types].xml 里指向已删除部件的 Override 去掉。</summary>
    private static void PruneContentTypeOverrides(ZipArchive zip, HashSet<string> keep)
    {
        var entry = zip.GetEntry("[Content_Types].xml");
        if (entry == null) return;

        XDocument xdoc;
        using (var stream = entry.Open())
        {
            xdoc = XDocument.Load(stream);
        }

        var root = xdoc.Root;
        if (root == null) return;

        var doomed = root.Elements()
            .Where(e => e.Name.LocalName == "Override")
            .Where(e =>
            {
                string? partName = (string?)e.Attribute("PartName");
                if (string.IsNullOrEmpty(partName)) return true;
                return !keep.Contains(Uri.UnescapeDataString(partName).TrimStart('/'));
            })
            .ToList();

        if (doomed.Count == 0) return;

        foreach (var e in doomed) e.Remove();

        string name = entry.FullName;
        entry.Delete();
        var created = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var output = created.Open();
        using var writer = new StreamWriter(output, new System.Text.UTF8Encoding(false));
        xdoc.Save(writer, SaveOptions.DisableFormatting);
    }

    /// <summary>读一个 .rels 文件里的内部目标路径（跳过 TargetMode="External"）。</summary>
    private static IEnumerable<string> ReadRelationshipTargets(ZipArchive zip, string relsPath, string baseDir)
    {
        var entry = zip.GetEntry(relsPath);
        if (entry == null) yield break;

        XDocument xdoc;
        using (var stream = entry.Open())
        {
            xdoc = XDocument.Load(stream);
        }

        if (xdoc.Root == null) yield break;

        foreach (var rel in xdoc.Root.Elements())
        {
            if (rel.Name.LocalName != "Relationship") continue;
            if (string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;

            string? target = (string?)rel.Attribute("Target");
            if (string.IsNullOrEmpty(target)) continue;

            yield return ResolvePartPath(baseDir, target);
        }
    }

    /// <summary>把关系里的相对路径解析成 zip 内的规范路径（处理 ../ 和 / 开头）。</summary>
    private static string ResolvePartPath(string baseDir, string target)
    {
        string text = Uri.UnescapeDataString(target).Replace('\\', '/');
        string combined = text.StartsWith("/", StringComparison.Ordinal)
            ? text.TrimStart('/')
            : (baseDir.Length == 0 ? text : baseDir + "/" + text);

        var segments = new List<string>();
        foreach (string seg in combined.Split('/'))
        {
            if (seg.Length == 0 || seg == ".") continue;
            if (seg == "..")
            {
                if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(seg);
        }
        return string.Join("/", segments);
    }

    private static string PartDirectory(string part)
    {
        int i = part.LastIndexOf('/');
        return i < 0 ? "" : part.Substring(0, i);
    }

    private static string RelationshipPartPath(string part)
    {
        string dir = PartDirectory(part);
        string file = dir.Length == 0 ? part : part.Substring(dir.Length + 1);
        return (dir.Length == 0 ? "_rels/" : dir + "/_rels/") + file + ".rels";
    }

    /// <summary>分析分节结构（只读），结果通过 log 回调输出。</summary>
    public static void AnalyzeSections(WordprocessingDocument doc, string filePath, Action<string> log)
    {
        log($"===================================");
        log($"正在分析文件: {Path.GetFileName(filePath)}");

        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null)
        {
            log("  [警告] 无法读取文档主体。");
            return;
        }

        var sections = ScanSections(doc);
        if (sections.Count == 0)
        {
            log("  [警告] 未能识别出任何节。");
            return;
        }

        int paragraphs = 0;
        int tables = 0;

        foreach (var s in sections)
        {
            paragraphs += s.Paragraphs;
            tables += s.Tables;

            string marks = s.HeaderRefs + s.FooterRefs > 0
                ? $"页眉 {s.HeaderRefs} / 页脚 {s.FooterRefs}"
                : "无页眉页脚引用";
            string content = s.HasContent ? "有内容" : "空白";

            log($"  -> 节 {s.Number}: 块 {s.FirstBlock}-{s.LastBlock}，段落 {s.Paragraphs}，表格 {s.Tables}，{marks}，{content}");
        }

        log($"  -------------------------------");
        log($"  统计结果:");
        log($"  - 节数量: {sections.Count}");
        log($"  - 段落数量: {paragraphs}");
        log($"  - 表格数量: {tables}");

        // 末尾连续空白节的个数（第 1 节永远留着，删到它就没正文了）
        int trailingBlank = 0;
        for (int i = sections.Count - 1; i >= 1 && !sections[i].HasContent; i--) trailingBlank++;

        if (sections.Count <= 1)
        {
            log($"  => 结论: 只有 1 个节，再删就没有正文了，程序会跳过。");
        }
        else if (trailingBlank == 0)
        {
            log($"  => 结论: 共 {sections.Count} 个节，最后一节还有内容，「删除末尾空白节」会跳过它。");
        }
        else
        {
            log($"  => 结论: 共 {sections.Count} 个节，末尾 {trailingBlank} 个是空白，「删除末尾空白节」会清掉这 {trailingBlank} 个。");
        }
    }

    /// <summary>计算一个文件的哈希值（小写十六进制）。</summary>
    public static string ComputeFileHash(string filePath, string algo)
    {
        using var fs = File.OpenRead(filePath);
        return ComputeStreamHash(fs, algo);
    }

    /// <summary>计算一个流的哈希值（小写十六进制）。</summary>
    public static string ComputeStreamHash(Stream stream, string algo)
    {
        stream.Position = 0;
        byte[] hash = NormalizeAlgorithm(algo) switch
        {
            "SHA1"   => SHA1.HashData(stream),
            "SHA256" => SHA256.HashData(stream),
            _        => MD5.HashData(stream),
        };
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }

    /// <summary>
    /// 找出容器中引用指定部件的所有关系 ID。
    /// 注意：SDK 3.x 中主文档图片关系可能不在 DataPartReferenceRelationships 里，
    /// 因此这里直接遍历 Parts（关系 ID → 部件）映射。
    /// </summary>
    private static IEnumerable<string> GetRelationshipIdsFor(OpenXmlPartContainer container, OpenXmlPart part)
    {
        foreach (var pair in container.Parts)
        {
            if (object.ReferenceEquals(pair.OpenXmlPart, part))
            {
                yield return pair.RelationshipId;
            }
        }

        string? id = container.GetIdOfPart(part);
        if (!string.IsNullOrEmpty(id))
        {
            yield return id;
        }
    }

    #region 水印 / 来源检测

    /// <summary>
    /// 「小于这个尺寸的图」判定为埋点/追踪图（单位 pt，1pt≈0.35mm）。正常插图不可能这么小，
    /// 常见的隐形追踪图写法是 width:.05pt;height:.05pt。
    /// </summary>
    private const double TinyImagePt = 2.0;

    /// <summary>
    /// 水印/来源检测。执行顺序是「先清水印，再问图片」：
    /// ① 扫文字和图形属性，判定有没有命中；
    /// ② 命中就先删水印对象和命中文字（图形水印、logo 图、埋点图、关键词文字）；
    /// ③ 再把**剩下的**图片拿去做 OCR，识别到关键词的图片单独删；
    /// ④ 最后收一遍孤儿媒体。
    /// 这样 OCR 不必去认那些马上要被删掉的水印 logo，判断的也是「清完水印后的图」。
    /// </summary>
    public static WatermarkReport RemoveWatermarks(WordprocessingDocument doc, IReadOnlyList<string> keywords,
                                                   bool dropAllImages, string? ocrDir, int ocrMaxImages,
                                                   Action<string>? log = null)
    {
        var report = new WatermarkReport();
        var kws = NormalizeKeywords(keywords);
        if (kws.Count == 0) return report;

        var parts = CollectContentParts(doc);
        if (parts.Count == 0) return report;

        // ① 只读探测：文字、VML 属性、图片替换文字
        CollectTextEvidence(parts, kws, report);

        // ② 命中就先清水印 + 命中文字。删全部图时图形清理没意义（后面整片删），单独走
        if (report.TextHit)
        {
            foreach (var part in parts)
            {
                var root = part.RootElement;
                if (root == null) continue;

                if (!dropAllImages) report.ShapesRemoved += RemoveWatermarkShapes(root, kws);
                report.KeywordsRemoved += RemoveKeywordText(root, kws);
            }
        }

        // ③ 再对剩下的图片做 OCR。已经确认命中又要把图全删时，识别没有意义
        var ocrHits = new Dictionary<OpenXmlPart, HashSet<string>>();
        if (!string.IsNullOrWhiteSpace(ocrDir) && !(dropAllImages && report.TextHit))
        {
            if (report.ShapesRemoved > 0 || report.KeywordsRemoved > 0)
            {
                log?.Invoke($"已先清水印对象 {report.ShapesRemoved} 处、命中文字 {report.KeywordsRemoved} 处，"
                            + "接下来只识别剩下的图片");
            }

            RunOcrOnImages(parts, ocrDir!, ocrMaxImages, kws, ocrHits, report, log);

            foreach (var part in parts)
            {
                var root = part.RootElement;
                if (root == null) continue;
                if (!ocrHits.TryGetValue(part, out var hitRelIds) || hitRelIds.Count == 0) continue;

                report.ShapesRemoved += RemoveImageReferences(root, hitRelIds);
            }
        }

        report.Hit = report.TextHit || ocrHits.Count > 0;
        if (!report.Hit) return report;

        if (dropAllImages)
        {
            int elements = 0;
            foreach (var part in parts) elements += RemoveAllImageElements(part);
            var (allFiles, allBytes) = RemoveOrphanMedia(doc);
            report.DroppedAllImages = true;
            report.ShapesRemoved = elements;
            report.ImagesRemoved = allFiles;
            if (allBytes > 0) log?.Invoke($"释放 {FormatSize(allBytes)}");
            return report;
        }

        // ④ 水印对象删掉后图片部件还挂着关系，交给已有的孤儿媒体清理收尾（它同时会删关系）
        if (report.ShapesRemoved > 0)
        {
            var (files, bytes) = RemoveOrphanMedia(doc);
            report.ImagesRemoved = files;
            if (files > 0) log?.Invoke($"顺带清掉不再被引用的媒体 {files} 个（{FormatSize(bytes)}）");
        }

        return report;
    }

    /// <summary>
    /// 只读探测：报告命中证据和「本来会删掉多少」，不改动文档。给 --dry-run 用。
    /// 和 <see cref="RemoveWatermarks"/> 共用同一套判定，所以干跑报的数字就是实跑会删的数字。
    /// </summary>
    public static WatermarkReport ScanWatermarks(WordprocessingDocument doc, IReadOnlyList<string> keywords)
    {
        var report = new WatermarkReport();
        var kws = NormalizeKeywords(keywords);
        if (kws.Count == 0) return report;

        var parts = CollectContentParts(doc);
        if (parts.Count == 0) return report;

        CollectTextEvidence(parts, kws, report);
        report.Hit = report.TextHit;

        foreach (var part in parts)
        {
            var root = part.RootElement;
            if (root == null) continue;

            report.ShapesRemoved += RemoveWatermarkShapes(root, kws, apply: false);
            report.KeywordsRemoved += RemoveKeywordText(root, kws, apply: false);

            report.ImageElements += root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Count()
                                    + root.Descendants().Count(d => d.LocalName == "imagedata");
        }

        return report;
    }

    /// <summary>
    /// 删掉所有部件里命中关键词的文字（正文、页眉、页脚、脚注、尾注都算）。分两种处理：
    /// ① 看不见的文字（白色 / ≤4pt）整段删——这类就是水印本体，只删关键词会留下
    ///    「（北京）股份有限公司」这种碎片；
    /// ② 看得见的文字只删关键词这几个字，同段其他文字保留（「本题由某某网整理」→「本题由整理」）。
    /// 跨 run 拆开写的（`某公`+`司`）也能删——按整段拼起来定位，再逐字回写。
    /// apply = false 时只统计不修改（干跑用）。
    /// 返回删掉（或将要删掉）的出现次数。
    /// </summary>
    private static int RemoveKeywordText(OpenXmlElement root, List<string> kws, bool apply = true)
    {
        int removed = 0;

        // ① 隐形文字整段删。记下这些 w:t，第 ② 步要跳过，否则干跑会把同一处文字数两遍
        var hiddenTexts = new HashSet<Text>();

        foreach (var run in root.Descendants<Run>().ToList())
        {
            string runText = string.Concat(run.Descendants<Text>().Select(t => t.Text));
            if (!ContainsKeyword(runText, kws)) continue;
            if (!IsInvisibleRun(run)) continue;

            removed += CountKeywordOccurrences(runText, kws);
            foreach (var t in run.Descendants<Text>()) hiddenTexts.Add(t);
            if (apply) run.Remove();
        }

        // ② 可见文字按字符删，只删关键词本身
        foreach (var paragraph in root.Descendants<Paragraph>().ToList())
        {
            var texts = paragraph.Descendants<Text>().Where(t => !hiddenTexts.Contains(t)).ToList();
            if (texts.Count == 0) continue;

            string all = string.Concat(texts.Select(t => t.Text));
            if (all.Length == 0 || !ContainsKeyword(all, kws)) continue;

            // 先整体标出要删的字符位置，再按 w:t 分组回写，避免改一处影响后面的下标
            var drop = new bool[all.Length];
            bool any = false;

            foreach (string k in kws)
            {
                int from = 0;
                while (from <= all.Length - k.Length)
                {
                    int idx = all.IndexOf(k, from, StringComparison.OrdinalIgnoreCase);
                    if (idx < 0) break;

                    for (int i = idx; i < idx + k.Length; i++) drop[i] = true;
                    any = true;
                    removed++;
                    from = idx + k.Length;
                }
            }

            if (!any || !apply) continue;

            int cursor = 0;
            foreach (var t in texts)
            {
                string s = t.Text ?? "";
                var kept = new System.Text.StringBuilder(s.Length);

                for (int i = 0; i < s.Length; i++)
                {
                    if (!drop[cursor + i]) kept.Append(s[i]);
                }
                cursor += s.Length;

                string fresh = kept.ToString();
                if (fresh == s) continue;

                if (fresh.Length == 0)
                {
                    // 整段文字都被删空：把空的 w:t 去掉，run 留着（可能还挂着制表符、图片等）
                    t.Remove();
                }
                else
                {
                    t.Text = fresh;
                    // 删完可能留下首尾空格，不加 preserve 会被 Word 吞掉
                    t.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        return removed;
    }

    /// <summary>这一段文字是不是「看不见」：白色字，或字号 ≤4pt（w:sz ≤ 8，单位半磅）。</summary>
    private static bool IsInvisibleRun(Run run)
    {
        var props = run.RunProperties;

        string? color = props?.Color?.Val?.Value;
        if (!string.IsNullOrEmpty(color) &&
            (color.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase) ||
             color.Equals("WHITE", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (props?.FontSize?.Val?.Value is string sz && int.TryParse(sz, out int halfPoints) && halfPoints <= 8)
        {
            return true;
        }

        // 字号也可能只写在段落标记的 rPr 里，取那里的值兜底
        if (run.Parent is Paragraph p)
        {
            var markSize = p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<FontSize>();
            if (markSize?.Val?.Value is string psz && int.TryParse(psz, out int ph) && ph <= 8)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>数一段文字里出现了几次关键词（只用于日志统计）。</summary>
    private static int CountKeywordOccurrences(string text, List<string> kws)
    {
        int n = 0;
        foreach (string k in kws)
        {
            int from = 0;
            while (from <= text.Length - k.Length)
            {
                int idx = text.IndexOf(k, from, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) break;
                n++;
                from = idx + k.Length;
            }
        }
        return n;
    }

    /// <summary>
    /// 删掉引用了指定图片关系 ID 的图形（OCR 命中关键词的那几张）。
    /// VML 的 v:shape 走 imagedata/@r:id，DrawingML 走 blip/@r:embed，mc:AlternateContent 连外壳一起删。
    /// </summary>
    private static int RemoveImageReferences(OpenXmlElement root, HashSet<string> relIds)
    {
        var toRemove = new List<OpenXmlElement>();

        foreach (var shape in root.Descendants().Where(d => d.LocalName == "shape").ToList())
        {
            var imgData = shape.Descendants().FirstOrDefault(d => d.LocalName == "imagedata");
            string? rid = imgData == null ? null : RelIdOf(imgData);
            if (rid != null && relIds.Contains(rid)) toRemove.Add(shape);
        }

        foreach (var blip in root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList())
        {
            string? rid = blip.Embed?.Value;
            if (rid == null || !relIds.Contains(rid)) continue;

            OpenXmlElement? target = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault();
            target ??= blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
            target ??= blip;
            if (target.Parent is Drawing) target = target.Parent;

            var alternate = target.Ancestors().FirstOrDefault(a => a.LocalName == "AlternateContent");
            if (alternate != null) target = alternate;

            toRemove.Add(target);
        }

        int removed = RemoveOutermost(toRemove);
        RemoveEmptyPicts(root);
        return removed;
    }

    private static List<string> NormalizeKeywords(IEnumerable<string> keywords)
        => keywords.Select(k => (k ?? "").Trim())
                   .Where(k => k.Length > 0)
                   .Distinct(StringComparer.OrdinalIgnoreCase)
                   .ToList();

    private static bool ContainsKeyword(string? text, List<string> keywords)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (string k in keywords)
        {
            if (text.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) return true;
        }
        return false;
    }

    /// <summary>扫描正文/页眉/页脚/脚注/尾注里的文字和 VML 属性，记录命中证据。</summary>
    private static void CollectTextEvidence(List<OpenXmlPart> parts, List<string> kws, WatermarkReport report)
    {
        foreach (var part in parts)
        {
            var root = part.RootElement;
            if (root == null) continue;

            string label = PartLabel(part);

            foreach (var t in root.Descendants<Text>())
            {
                if (!ContainsKeyword(t.Text, kws)) continue;
                report.TextHit = true;
                report.AddEvidence($"{label}文字「{Shorten(t.Text)}」");
            }

            // 艺术字水印的文字不在 w:t 里，而是在 v:shape/@alt、v:textpath/@string、v:imagedata/@o:title 这些属性里
            foreach (var el in root.Descendants())
            {
                foreach (var attr in el.GetAttributes())
                {
                    // alt/string/title 是 VML 艺术字水印，descr 是 DrawingML 图片的「替换文字」
                    if (attr.LocalName is not ("string" or "alt" or "title" or "descr")) continue;
                    if (!ContainsKeyword(attr.Value, kws)) continue;
                    report.TextHit = true;
                    report.AddEvidence($"{label}{el.LocalName}/@{attr.LocalName}「{Shorten(attr.Value)}」");
                }
            }
        }
    }

    /// <summary>
    /// 删掉一个部件里的水印图形：命中关键词的 VML 形状/成组水印、尺寸小到看不见的埋点图、
    /// 以及 descr 里带关键词的 DrawingML logo 图。返回删掉的元素个数。
    /// OCR 命中的图片走 <see cref="RemoveImageReferences"/>，不在这里处理。
    /// </summary>
    private static int RemoveWatermarkShapes(OpenXmlElement root, List<string> kws, bool apply = true)
    {
        var toRemove = new List<OpenXmlElement>();

        foreach (var shape in root.Descendants().Where(d => d.LocalName is "shape" or "group").ToList())
        {
            if (IsWatermarkShape(shape, kws) || (shape.LocalName == "shape" && IsTinyVmlShape(shape)))
            {
                toRemove.Add(shape);
            }
        }

        foreach (var blip in root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList())
        {
            bool hit = IsTinyDrawing(blip) || IsWatermarkPicture(blip, kws);
            if (!hit) continue;

            OpenXmlElement? target = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault();
            target ??= blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
            target ??= blip;
            if (target.Parent is Drawing) target = target.Parent;

            // mc:AlternateContent 里 Choice（DrawingML）和 Fallback（VML）是同一张图的两个版本，
            // 只删其中一个会留下空壳，连外壳一起删
            var alternate = target.Ancestors().FirstOrDefault(a => a.LocalName == "AlternateContent");
            if (alternate != null) target = alternate;

            toRemove.Add(target);
        }

        if (!apply)
        {
            // 干跑：只数「最外层会被删掉的元素」，和实跑的口径一致（父元素已算进去就不再算子元素）
            var set = new HashSet<OpenXmlElement>(toRemove);
            return set.Count(e => e.Parent != null && !e.Ancestors().Any(set.Contains));
        }

        int removed = RemoveOutermost(toRemove);

        // 形状删完后，只剩 shapetype 空壳的 w:pict 一并去掉，别留无用结构
        RemoveEmptyPicts(root);

        return removed;
    }

    /// <summary>
    /// 清掉已经没有图形可画的空壳：先由内向外删空 v:group，再删只剩 shapetype 的 w:pict。
    /// 成组水印常写成 v:group 包一个 v:shape，只删 shape 会留下带 alt 的空组。
    /// </summary>
    private static void RemoveEmptyPicts(OpenXmlElement root)
    {
        bool again = true;
        while (again)
        {
            again = false;

            foreach (var group in root.Descendants().Where(d => d.LocalName == "group").ToList())
            {
                if (group.ChildElements.Count > 0) continue;
                group.Remove();
                again = true;
            }

            foreach (var pict in root.Descendants().Where(d => d.LocalName == "pict").ToList())
            {
                bool hasDrawing = pict.Descendants().Any(d => d.LocalName is "shape" or "imagedata" or "group");
                if (hasDrawing) continue;
                pict.Remove();
                again = true;
            }
        }
    }

    /// <summary>这个 VML 形状是不是水印：id 里带 watermark，或 alt / 内嵌 textpath 的文字含关键词。</summary>
    private static bool IsWatermarkShape(OpenXmlElement shape, List<string> kws)
    {
        string id = LocalAttr(shape, "id") ?? "";
        if (id.IndexOf("watermark", StringComparison.OrdinalIgnoreCase) >= 0) return true;

        foreach (string name in new[] { "id", "alt", "title", "string" })
        {
            if (ContainsKeyword(LocalAttr(shape, name), kws)) return true;
        }

        foreach (var tp in shape.Descendants().Where(d => d.LocalName == "textpath"))
        {
            if (ContainsKeyword(LocalAttr(tp, "string"), kws)) return true;
        }

        return false;
    }

    /// <summary>VML 形状的 width/height 都小到看不见（埋点图的典型写法 0.05pt / 0.75pt）。</summary>
    private static bool IsTinyVmlShape(OpenXmlElement shape)
    {
        string style = LocalAttr(shape, "style") ?? "";
        double? w = LengthPtFromStyle(style, "width");
        double? h = LengthPtFromStyle(style, "height");
        return w is > 0 and <= TinyImagePt && h is > 0 and <= TinyImagePt;
    }

    /// <summary>
    /// DrawingML 图片是不是水印：来源站点的 logo 图常把品牌名写进 wp:docPr/@descr（图片替代文字），
    /// 或者 pic:cNvPr/@descr。这个属性 Word 界面上叫「替换文字」，用户看不见但它跟着图走。
    /// </summary>
    private static bool IsWatermarkPicture(OpenXmlElement blip, List<string> kws)
    {
        var frame = blip.Ancestors().FirstOrDefault(a => a.LocalName is "inline" or "anchor");
        if (frame == null) return false;

        foreach (var el in frame.Descendants().Prepend(frame))
        {
            if (el.LocalName is not ("docPr" or "cNvPr")) continue;
            foreach (string name in new[] { "descr", "name", "title" })
            {
                if (ContainsKeyword(LocalAttr(el, name), kws)) return true;
            }
        }

        return false;
    }

    /// <summary>DrawingML 图片的显示尺寸小到看不见（EMU，1pt = 12700 EMU）。</summary>
    private static bool IsTinyDrawing(OpenXmlElement blip)
    {
        var extent = blip.Ancestors().FirstOrDefault(a => a.LocalName == "extent");
        if (extent == null) return false;

        if (!long.TryParse(LocalAttr(extent, "cx"), out long cx)) return false;
        if (!long.TryParse(LocalAttr(extent, "cy"), out long cy)) return false;

        long limit = (long)(TinyImagePt * 12700);
        return cx > 0 && cy > 0 && cx <= limit && cy <= limit;
    }

    /// <summary>从 VML 的 style 串里取长度并换算成 pt，例如 "width:.05pt"。</summary>
    private static double? LengthPtFromStyle(string style, string key)
    {
        if (string.IsNullOrEmpty(style)) return null;

        var m = Regex.Match(style,
            @"(?:^|;)\s*" + Regex.Escape(key) + @"\s*:\s*(?<v>[0-9]*\.?[0-9]+)\s*(?<u>pt|in|cm|mm|px)?",
            RegexOptions.IgnoreCase);

        if (!m.Success) return null;
        if (!double.TryParse(m.Groups["v"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
            return null;

        string unit = m.Groups["u"].Success ? m.Groups["u"].Value.ToLowerInvariant() : "pt";
        return unit switch
        {
            "in" => v * 72.0,
            "cm" => v * 72.0 / 2.54,
            "mm" => v * 72.0 / 25.4,
            "px" => v * 72.0 / 96.0,
            _ => v,
        };
    }

    /// <summary>删掉一个部件里所有图片元素（DrawingML 的 blip + VML 的 imagedata），媒体部件交给孤儿清理。</summary>
    private static int RemoveAllImageElements(OpenXmlPart part)
    {
        var root = part.RootElement;
        if (root == null) return 0;

        var toRemove = new List<OpenXmlElement>();

        foreach (var blip in root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
        {
            OpenXmlElement? target = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault();
            target ??= blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
            target ??= blip;
            if (target.Parent is Drawing) target = target.Parent;
            toRemove.Add(target);
        }

        foreach (var imgData in root.Descendants().Where(d => d.LocalName == "imagedata").ToList())
        {
            var shape = imgData.Ancestors().FirstOrDefault(a => a.LocalName == "shape");
            toRemove.Add(shape ?? (OpenXmlElement)imgData);
        }

        int removed = RemoveOutermost(toRemove);

        RemoveEmptyPicts(root);

        return removed;
    }

    /// <summary>
    /// 只保留最外层元素再删。同一个图形会同时命中祖先（w:drawing）和它内部的后代（VML shape），
    /// 对已脱离树的元素调用 Remove() 会抛异常，整个文件就处理失败了。
    /// </summary>
    private static int RemoveOutermost(IEnumerable<OpenXmlElement> elements)
    {
        var set = new HashSet<OpenXmlElement>(elements);
        if (set.Count == 0) return 0;

        var outermost = set.Where(e => e.Parent != null && !e.Ancestors().Any(set.Contains)).ToList();
        foreach (var e in outermost) e.Remove();
        return outermost.Count;
    }

    private static string? LocalAttr(OpenXmlElement element, string localName)
        => element.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value;

    private static string? RelIdOf(OpenXmlElement element)
        => element.GetAttributes().FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == RelNs).Value;

    private static string PartLabel(OpenXmlPart part) => part switch
    {
        MainDocumentPart => "正文",
        HeaderPart => "页眉",
        FooterPart => "页脚",
        FootnotesPart => "脚注",
        EndnotesPart => "尾注",
        _ => part.GetType().Name,
    };

    private static string Shorten(string? text, int max = 24)
    {
        string s = (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
        return s.Length <= max ? s : s.Substring(0, max) + "…";
    }

    /// <summary>
    /// 把文档里的位图抽到临时目录，一次进程跑完 OCR，把命中关键词的图片记进 ocrHits（部件 → 关系 ID）。
    /// 只做位图：emf/wmf/svg 这些 ppocr 读不了。小于 2KB 的多半是 1x1 埋点图，识别没有意义。
    /// </summary>
    private static void RunOcrOnImages(List<OpenXmlPart> parts, string ocrDir, int maxImages, List<string> kws,
                                       Dictionary<OpenXmlPart, HashSet<string>> ocrHits, WatermarkReport report,
                                       Action<string>? log)
    {
        string? home = LocalOcr.ResolveHome(ocrDir);
        if (home == null || !File.Exists(Path.Combine(home, "ppocr.exe")))
        {
            report.OcrNote = $"找不到 OCR 程序（{ocrDir}），已跳过图片识别";
            log?.Invoke(report.OcrNote);
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "DocxToolbox-ocr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            var map = new Dictionary<string, (OpenXmlPart Part, string RelId)>(StringComparer.OrdinalIgnoreCase);
            int index = 0, skipped = 0;

            foreach (var part in parts)
            {
                foreach (var pair in part.Parts.ToList())
                {
                    if (pair.OpenXmlPart is not ImagePart img) continue;

                    string ext = Path.GetExtension(img.Uri.ToString()).ToLowerInvariant();
                    if (ext is not (".png" or ".jpg" or ".jpeg" or ".bmp" or ".tif" or ".tiff" or ".gif"))
                    {
                        skipped++;
                        continue;
                    }

                    long size;
                    try
                    {
                        using var probe = img.GetStream(FileMode.Open, FileAccess.Read);
                        size = probe.Length;
                    }
                    catch { skipped++; continue; }

                    if (size < 2048) { skipped++; continue; }
                    if (maxImages > 0 && index >= maxImages) { skipped++; continue; }

                    string name = $"img{index:0000}{ext}";
                    try
                    {
                        using var src = img.GetStream(FileMode.Open, FileAccess.Read);
                        using var dst = File.Create(Path.Combine(tempDir, name));
                        src.CopyTo(dst);
                    }
                    catch { skipped++; continue; }

                    map[name] = (part, pair.RelationshipId);
                    index++;
                }
            }

            if (map.Count == 0)
            {
                report.OcrNote = "没有可识别的位图";
                log?.Invoke(report.OcrNote);
                return;
            }

            var started = DateTime.Now;
            var timeout = TimeSpan.FromSeconds(90 + map.Count * 4);
            var results = LocalOcr.Recognize(home, tempDir, timeout, log);
            report.OcrScanned = map.Count;

            foreach (var kv in results)
            {
                if (!ContainsKeyword(kv.Value, kws)) continue;
                if (!map.TryGetValue(kv.Key, out var target)) continue;

                if (!ocrHits.TryGetValue(target.Part, out var set))
                    ocrHits[target.Part] = set = new HashSet<string>(StringComparer.Ordinal);

                set.Add(target.RelId);
                report.AddEvidence($"图片 {kv.Key} 识别出「{Shorten(kv.Value)}」");
            }

            log?.Invoke($"OCR 识别 {map.Count} 张图片，用时 {(int)(DateTime.Now - started).TotalSeconds} 秒"
                        + (skipped > 0 ? $"，跳过 {skipped} 张（过小或非位图）" : ""));
        }
        finally
        {
            try { Directory.Delete(tempDir, true); } catch { /* 临时目录删不掉不影响结果 */ }
        }
    }

    #endregion

    public static int RemoveImagesByHashes(WordprocessingDocument doc, IEnumerable<string> targetHashes,
                                           string algo, Action<string>? log = null)
    {
        int totalDeleted = 0;

        var targets = new HashSet<string>(
            targetHashes.Where(h => !string.IsNullOrWhiteSpace(h)).Select(h => h.Trim()),
            StringComparer.OrdinalIgnoreCase);
        if (targets.Count == 0) return 0;

        var mainPart = doc.MainDocumentPart;
        if (mainPart == null) return 0;

        var partsToScan = new List<OpenXmlPart>();
        partsToScan.Add(mainPart);
        partsToScan.AddRange(mainPart.HeaderParts);
        partsToScan.AddRange(mainPart.FooterParts);
        if (mainPart.FootnotesPart != null) partsToScan.Add(mainPart.FootnotesPart);
        if (mainPart.EndnotesPart  != null) partsToScan.Add(mainPart.EndnotesPart);

        foreach (var part in partsToScan)
        {
            var matched = new List<ImagePart>();
            var relIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var ip in part.GetPartsOfType<ImagePart>())
            {
                string hash;
                try
                {
                    using var stream = ip.GetStream(FileMode.Open, FileAccess.Read);
                    hash = ComputeStreamHash(stream, algo);
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[警告] 读取图片流失败 ({part.GetType().Name}): {ex.Message}");
                    continue;
                }

                if (targets.Contains(hash))
                {
                    matched.Add(ip);
                    foreach (string rid in GetRelationshipIdsFor(part, ip))
                    {
                        relIds.Add(rid);
                    }
                }
            }

            if (matched.Count == 0) continue;

            log?.Invoke($"在 {part.GetType().Name} 中发现 {matched.Count} 张匹配图片，rId: {string.Join(", ", relIds)}");

            var root = part.RootElement;
            if (root != null)
            {
                var toRemove = new List<OpenXmlElement>();

                foreach (var blip in root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList())
                {
                    if (blip.Embed == null ||
                        string.IsNullOrEmpty(blip.Embed.Value) ||
                        !relIds.Contains(blip.Embed.Value))
                    {
                        continue;
                    }

                    // 整块移除承载图片的 drawing（Inline/Anchor）：只删 Inline 会留下非法的空 <w:drawing/>
                    OpenXmlElement? target = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault();
                    target ??= blip.Ancestors<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
                    target ??= blip;

                    if (target.Parent is Drawing)
                    {
                        target = target.Parent;
                    }

                    toRemove.Add(target);
                }

                foreach (var imgData in root.Descendants().Where(d => d.LocalName == "imagedata").ToList())
                {
                    var idAttr = imgData.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    if (!string.IsNullOrEmpty(idAttr.Value) && relIds.Contains(idAttr.Value))
                    {
                        var shape = imgData.Ancestors().FirstOrDefault(a => a.LocalName == "shape");
                        if (shape != null)
                            toRemove.Add(shape);
                        else
                            toRemove.Add(imgData);
                    }
                }

                // 同一个图形可能被同时收进祖先（w:drawing）和它内部的后代（VML shape），
                // 只保留最外层，否则对已脱离文档树的元素调用 Remove() 会抛异常、整个文件处理失败
                var removeSet = new HashSet<OpenXmlElement>(toRemove);
                var outermost = removeSet
                    .Where(e => e.Parent != null && !e.Ancestors().Any(removeSet.Contains))
                    .ToList();

                foreach (var elem in outermost)
                {
                    log?.Invoke($"  删除元素: {elem.GetType().Name}");
                    elem.Remove();
                }
            }

            foreach (var ip in matched)
            {
                // DeletePart 会一并移除容器中指向该图片的所有关系
                try
                {
                    part.DeletePart(ip);
                    totalDeleted++;
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[警告] 删除 ImagePart 失败: {ex.Message}");
                }
            }
        }

        return totalDeleted;
    }
}
