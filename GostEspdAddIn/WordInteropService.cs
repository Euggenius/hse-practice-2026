using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace GostEspdAddIn.Services
{
    public sealed class WordInteropService
    {
        private const float IndentTolerance = 2.0f;

        private static readonly Regex ManualMarkerRegex =
            new Regex(
                @"^(?<indent>[ \t]*)" +
                @"(?<marker>" +
                    @"[•●○■▪▫◦]" +
                    @"|[-–—]" +
                    @"|\*" +
                    @"|\d{1,4}[\.\)]" +
                    @"|[А-Яа-яЁёA-Za-z]{1,3}[\.\)]" +
                @")" +
                @"(?<separator>[ \t]+)",
                RegexOptions.Compiled);

        private sealed class ParagraphInfo
        {
            public Paragraph Paragraph { get; set; }

            public string RawText { get; set; }

            public string Content { get; set; }

            public bool HasMarker { get; set; }

            public bool HasManualMarker { get; set; }

            public bool IsWordList { get; set; }

            public int? WordListLevel { get; set; }

            public int Level { get; set; }

            public float LeftIndent { get; set; }

            public float FirstLineIndent { get; set; }

            public int LeadingTabCount { get; set; }

            public int LeadingSpaceCount { get; set; }

            public float OriginalIndent { get; set; }
        }

        public void FormatSelectedList(Application wordApp)
        {
            if (wordApp == null)
                throw new ArgumentNullException(nameof(wordApp));

            Selection selection =
                wordApp.Selection;

            if (selection == null ||
                selection.Paragraphs == null ||
                selection.Paragraphs.Count == 0 ||
                string.IsNullOrWhiteSpace(selection.Text))
            {
                System.Windows.Forms.MessageBox.Show(
                    "Выделите текст для форматирования!",
                    "Внимание",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

                return;
            }

            bool oldScreenUpdating =
                wordApp.ScreenUpdating;

            try
            {
                wordApp.ScreenUpdating = false;

                UndoRecord undo =
                    wordApp.UndoRecord;

                undo.StartCustomRecord(
                    "Форматирование перечисления");

                try
                {
                    List<ParagraphInfo> items =
                        ReadParagraphs(selection);

                    if (items.Count == 0)
                        return;

                    AnalyzeStructure(items);

                    FormatParagraphs(items);
                }
                finally
                {
                    undo.EndCustomRecord();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.ToString(),
                    "Ошибка форматирования",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating =
                    oldScreenUpdating;

                wordApp.ScreenRefresh();
            }
        }

        private List<ParagraphInfo> ReadParagraphs(
            Selection selection)
        {
            var result =
                new List<ParagraphInfo>();

            foreach (Paragraph paragraph
                     in selection.Paragraphs)
            {
                if (paragraph == null)
                    continue;

                Range range =
                    paragraph.Range.Duplicate;

                string rawText =
                    RemoveParagraphEnd(
                        range.Text);

                if (string.IsNullOrWhiteSpace(rawText))
                    continue;

                var item =
                    new ParagraphInfo
                    {
                        Paragraph = paragraph,
                        RawText = rawText,

                        LeftIndent =
                            paragraph.Format.LeftIndent,

                        FirstLineIndent =
                            paragraph.Format.FirstLineIndent
                    };

                int position = 0;

                while (position < rawText.Length)
                {
                    if (rawText[position] == '\t')
                    {
                        item.LeadingTabCount++;
                        position++;
                    }
                    else if (rawText[position] == ' ')
                    {
                        item.LeadingSpaceCount++;
                        position++;
                    }
                    else
                    {
                        break;
                    }
                }

                try
                {
                    item.IsWordList =
                        paragraph.Range.ListFormat.ListType !=
                        WdListType.wdListNoNumbering;

                    if (item.IsWordList)
                    {
                        item.WordListLevel =
                            paragraph.Range
                                .ListFormat
                                .ListLevelNumber;
                    }
                }
                catch
                {
                    item.IsWordList = false;
                    item.WordListLevel = null;
                }

                Match match =
                    ManualMarkerRegex.Match(
                        rawText);

                item.HasManualMarker =
                    match.Success;

                item.HasMarker =
                    item.HasManualMarker ||
                    item.IsWordList;

                if (match.Success)
                {
                    item.Content =
                        rawText.Substring(
                            match.Length);

                    item.Content =
                        RemoveLeadingWhitespace(
                            item.Content);
                }
                else
                {
                    item.Content =
                        RemoveLeadingWhitespace(
                            rawText);
                }

                item.OriginalIndent =
                    CalculateOriginalIndent(
                        item);

                result.Add(item);
            }

            return result;
        }

        private float CalculateOriginalIndent(
            ParagraphInfo item)
        {
            float indent =
                item.LeftIndent +
                Math.Max(
                    0,
                    item.FirstLineIndent);

            return indent;
        }

        private void AnalyzeStructure(
            List<ParagraphInfo> items)
        {
            if (items.Count == 0)
                return;

            List<float> indentLevels =
                items
                    .Where(x =>
                        x.HasMarker &&
                        x.LeadingTabCount == 0)
                    .Select(x =>
                        x.OriginalIndent)
                    .DistinctWithTolerance(
                        IndentTolerance)
                    .OrderBy(x => x)
                    .ToList();

            foreach (ParagraphInfo item in items)
            {
                if (!item.HasMarker)
                {
                    item.Level = 0;
                    continue;
                }

                if (item.IsWordList &&
                    item.WordListLevel.HasValue &&
                    item.WordListLevel.Value > 0)
                {
                    item.Level =
                        item.WordListLevel.Value;

                    continue;
                }

                if (item.LeadingTabCount > 0)
                {
                    item.Level =
                        item.LeadingTabCount + 1;

                    continue;
                }

                if (indentLevels.Count > 0)
                {
                    item.Level =
                        FindNearestLevel(
                            item.OriginalIndent,
                            indentLevels);
                }
                else
                {
                    item.Level = 1;
                }

                if (item.Level < 1)
                    item.Level = 1;
            }

            int previousLevel = 1;

            foreach (ParagraphInfo item in items)
            {
                if (!item.HasMarker)
                    continue;

                if (item.Level >
                    previousLevel + 1)
                {
                    item.Level =
                        previousLevel + 1;
                }

                previousLevel =
                    item.Level;
            }
        }

        private int FindNearestLevel(
            float value,
            List<float> levels)
        {
            if (levels.Count == 0)
                return 1;

            int bestIndex = 0;

            float bestDistance =
                Math.Abs(
                    levels[0] - value);

            for (int i = 1;
                 i < levels.Count;
                 i++)
            {
                float distance =
                    Math.Abs(
                        levels[i] - value);

                if (distance < bestDistance)
                {
                    bestIndex = i;
                    bestDistance = distance;
                }
            }

            return bestIndex + 1;
        }

        private void FormatParagraphs(
            List<ParagraphInfo> items)
        {
            for (int i = 0;
                 i < items.Count;
                 i++)
            {
                ParagraphInfo item =
                    items[i];

                if (!item.HasMarker)
                {
                    if (IsIntroductoryText(
                            items,
                            i))
                    {
                        EnsureColonAtEnd(
                            item.Paragraph);
                    }

                    continue;
                }

                string marker =
                    BuildMarker(
                        items,
                        i);

                string punctuation =
                    DeterminePunctuation(
                        items,
                        i);

                NormalizeParagraph(
                    item,
                    marker,
                    punctuation);
            }
        }

        private string BuildMarker(
            List<ParagraphInfo> items,
            int index)
        {
            ParagraphInfo current =
                items[index];

            if (current.Level == 1)
                return "– ";

            int parentLevel =
                current.Level - 1;

            int parentIndex = -1;

            for (int i = index - 1;
                 i >= 0;
                 i--)
            {
                if (!items[i].HasMarker)
                    continue;

                if (items[i].Level ==
                    parentLevel)
                {
                    parentIndex = i;
                    break;
                }

                if (items[i].Level <
                    parentLevel)
                {
                    break;
                }
            }

            int number = 1;

            if (parentIndex >= 0)
            {
                for (int i =
                         parentIndex + 1;
                     i < index;
                     i++)
                {
                    if (!items[i].HasMarker)
                        continue;

                    if (items[i].Level <
                        current.Level)
                    {
                        break;
                    }

                    if (items[i].Level ==
                        current.Level)
                    {
                        number++;
                    }
                }
            }
            else
            {
                for (int i = 0;
                     i < index;
                     i++)
                {
                    if (items[i].HasMarker &&
                        items[i].Level ==
                        current.Level)
                    {
                        number++;
                    }
                }
            }

            return number + ") ";
        }

        private string DeterminePunctuation(
            List<ParagraphInfo> items,
            int index)
        {
            ParagraphInfo current =
                items[index];

            if (IsLastListItem(
                    items,
                    index))
            {
                return ".";
            }

            if (index + 1 < items.Count)
            {
                ParagraphInfo next =
                    items[index + 1];

                if (next.HasMarker &&
                    next.Level >
                    current.Level)
                {
                    return ":";
                }
            }

            return ";";
        }

        private bool IsLastListItem(
            List<ParagraphInfo> items,
            int index)
        {
            for (int i = index + 1;
                 i < items.Count;
                 i++)
            {
                if (items[i].HasMarker)
                    return false;
            }

            return true;
        }

        private void NormalizeParagraph(
            ParagraphInfo item,
            string marker,
            string punctuation)
        {
            Paragraph paragraph =
                item.Paragraph;

            if (item.IsWordList)
            {
                try
                {
                    paragraph.Range
                        .ListFormat
                        .RemoveNumbers();
                }
                catch
                {
                }
            }

            Range range =
                paragraph.Range.Duplicate;

            string text =
                RemoveParagraphEnd(
                    range.Text);

            Match oldMarker =
                ManualMarkerRegex.Match(
                    text);

            if (oldMarker.Success)
            {
                Range markerRange =
                    paragraph.Range.Duplicate;

                markerRange.Start =
                    paragraph.Range.Start;

                markerRange.End =
                    paragraph.Range.Start +
                    oldMarker.Length;

                markerRange.Delete();
            }
            else
            {
                Match whitespace =
                    Regex.Match(
                        text,
                        @"^[ \t]+");

                if (whitespace.Success)
                {
                    Range leading =
                        paragraph.Range.Duplicate;

                    leading.Start =
                        paragraph.Range.Start;

                    leading.End =
                        paragraph.Range.Start +
                        whitespace.Length;

                    leading.Delete();
                }
            }

            Range insertion =
                paragraph.Range.Duplicate;

            insertion.Collapse(
                WdCollapseDirection
                    .wdCollapseStart);

            insertion.InsertBefore(
                marker);

            Range newMarker =
                paragraph.Range.Duplicate;

            newMarker.Start =
                paragraph.Range.Start;

            newMarker.End =
                Math.Min(
                    paragraph.Range.Start +
                    marker.Length,
                    paragraph.Range.End - 1);

            try
            {
                newMarker.Font.Reset();
            }
            catch
            {
            }

            RemoveEndingPunctuation(
                paragraph,
                marker.Length);

            InsertEndingPunctuation(
                paragraph,
                punctuation);

            NormalizeFirstLetter(
                paragraph,
                marker);

            ApplyGostParagraphFormatting(
                paragraph,
                item.Level);
        }

        private void RemoveEndingPunctuation(
            Paragraph paragraph,
            int markerLength)
        {
            Range range =
                paragraph.Range.Duplicate;

            string text =
                RemoveParagraphEnd(
                    range.Text);

            if (text.Length <= markerLength)
                return;

            int end =
                text.Length;

            while (end > markerLength &&
                   char.IsWhiteSpace(
                       text[end - 1]))
            {
                end--;
            }

            while (end > markerLength &&
                   IsEndingPunctuation(
                       text[end - 1]))
            {
                end--;
            }

            if (end >= text.Length)
                return;

            Range tail =
                paragraph.Range.Duplicate;

            tail.Start =
                paragraph.Range.Start +
                end;

            tail.End =
                paragraph.Range.Start +
                text.Length;

            tail.Delete();
        }

        private void InsertEndingPunctuation(
            Paragraph paragraph,
            string punctuation)
        {
            Range range =
                paragraph.Range.Duplicate;

            range.Start =
                paragraph.Range.End - 1;

            range.End =
                range.Start;

            range.InsertBefore(
                punctuation);
        }

        private bool IsEndingPunctuation(
            char c)
        {
            return c == '.' ||
                   c == ',' ||
                   c == ';' ||
                   c == ':' ||
                   c == '!' ||
                   c == '?' ||
                   c == '…';
        }

        private void NormalizeFirstLetter(
            Paragraph paragraph,
            string marker)
        {
            Range range =
                paragraph.Range.Duplicate;

            string text =
                RemoveParagraphEnd(
                    range.Text);

            if (text.Length <= marker.Length)
                return;

            int position =
                marker.Length;

            while (position < text.Length &&
                   char.IsWhiteSpace(
                       text[position]))
            {
                position++;
            }

            if (position >= text.Length)
                return;

            char first =
                text[position];

            if (!char.IsLetter(first))
                return;

            if (!char.IsUpper(first))
                return;

            int uppercaseCount = 0;

            for (int i = position;
                 i < text.Length &&
                 i < position + 8;
                 i++)
            {
                if (char.IsUpper(text[i]))
                {
                    uppercaseCount++;
                }
                else if (char.IsLetter(
                             text[i]))
                {
                    break;
                }
            }

            if (uppercaseCount >= 2)
                return;

            if (position + 1 < text.Length &&
                char.IsUpper(
                    text[position + 1]))
            {
                return;
            }

            Range firstChar =
                paragraph.Range.Duplicate;

            firstChar.Start =
                paragraph.Range.Start +
                position;

            firstChar.End =
                firstChar.Start + 1;

            try
            {
                firstChar.Case =
                    WdCharacterCase
                        .wdLowerCase;
            }
            catch
            {
            }
        }

        private bool IsIntroductoryText(
            List<ParagraphInfo> items,
            int index)
        {
            if (items[index].HasMarker)
                return false;

            for (int i = 0;
                 i < index;
                 i++)
            {
                if (items[i].HasMarker)
                    return false;
            }

            for (int i = index + 1;
                 i < items.Count;
                 i++)
            {
                if (items[i].HasMarker)
                    return true;

                if (!string.IsNullOrWhiteSpace(
                        items[i].RawText))
                {
                    return false;
                }
            }

            return false;
        }

        private void EnsureColonAtEnd(
            Paragraph paragraph)
        {
            Range range =
                paragraph.Range.Duplicate;

            string text =
                RemoveParagraphEnd(
                    range.Text);

            int end =
                text.Length;

            while (end > 0 &&
                   char.IsWhiteSpace(
                       text[end - 1]))
            {
                end--;
            }

            if (end == 0)
                return;

            if (text[end - 1] == ':')
                return;

            while (end > 0 &&
                   IsEndingPunctuation(
                       text[end - 1]))
            {
                end--;
            }

            if (end < text.Length)
            {
                Range tail =
                    paragraph.Range.Duplicate;

                tail.Start =
                    paragraph.Range.Start +
                    end;

                tail.End =
                    paragraph.Range.End - 1;

                tail.Delete();
            }

            Range insert =
                paragraph.Range.Duplicate;

            insert.Start =
                paragraph.Range.End - 1;

            insert.End =
                insert.Start;

            insert.InsertBefore(":");
        }

        private void ApplyGostParagraphFormatting(
            Paragraph paragraph,
            int level)
        {
            ParagraphFormat format =
                paragraph.Format;

            format.Alignment =
                WdParagraphAlignment
                    .wdAlignParagraphJustify;

            format.LineSpacingRule =
                WdLineSpacing
                    .wdLineSpace1pt5;

            format.SpaceBefore = 0;
            format.SpaceAfter = 0;

            float baseIndent =
                paragraph.Application
                    .CentimetersToPoints(
                        1.25f);

            float levelStep =
                paragraph.Application
                    .CentimetersToPoints(
                        0.625f);

            format.LeftIndent =
                baseIndent +
                levelStep *
                Math.Max(
                    0,
                    level - 1);

            format.FirstLineIndent = 0;
        }

        private string RemoveParagraphEnd(
            string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            return text.TrimEnd(
                '\r',
                '\n',
                '\a');
        }

        private string RemoveLeadingWhitespace(
            string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            int index = 0;

            while (index < text.Length &&
                   (text[index] == ' ' ||
                    text[index] == '\t'))
            {
                index++;
            }

            return text.Substring(index);
        }
    }
}