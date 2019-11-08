using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace ComObject.Test
{
    public class WordInteropTests
        : IDisposable
    {
        [Fact]
        public void Save_document()
        {
            var docx = GetTempFile(".pdf");

            _document.SaveAs(docx);

            Assert.True(File.Exists(docx));
        }

        [Fact]
        public void Save_document_as_doc()
        {
            const int wdFormatDocument = 0;

            var doc = GetTempFile(".doc");

            _document.SaveAs(doc, wdFormatDocument);

            Assert.True(File.Exists(doc));
        }

        [Fact]
        public void Save_document_as_pdf()
        {
            const int wdFormatPDF = 17;

            var pdf = GetTempFile(".pdf");

            _document.SaveAs(pdf, wdFormatPDF);

            Assert.True(File.Exists(pdf));
        }

        [Fact]
        public void Add_paragraph()
        {
            var paragraph = _document.Paragraphs.Add();

            const string text = "Pagragraph 1";

            paragraph.Range.Text = text;

            Assert.Equal(_document.Paragraphs[1].Range.Text.Trim(), text);
        }

        [Theory]
        [InlineData("Text", "TextReplace")]
        public void Replace_text(
            string search,
            string replace)
        {
            const int wdFindContinue = 1;
            const int wdReplaceAll = 2;

            var paragraph = _document.Paragraphs.Add();

            paragraph.Range.Text = search;

            _application.Selection.Find.ClearFormatting();
            _application.Selection.Find.Replacement.ClearFormatting();

            var find = _application.Selection.Find;

            find.Text = search;
            find.Replacement.Text = replace;
            find.Forward = true;
            find.Wrap = wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            _application.Selection.Find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, wdReplaceAll);

            Assert.Equal(replace, _document.Paragraphs[1].Range.Text.Trim());
        }

        public WordInteropTests()
        {
            _application = new ComObject("Word.Application");

            _application.Visible = false;

            _document = _application.Documents.Add();
        }

        private string GetTempFile(
            string extension)
        {
            var file = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetTempFileName(), extension));

            _files.Add(file);

            return file;
        }

        public void Dispose()
        {
            _document.Close(false);
            _document.Dispose();

            _application.Quit(false);
            _application.Dispose();

            foreach (var file in _files.Where(File.Exists))
            {
                File.Delete(file);
            }
        }

        private readonly dynamic _application;
        private readonly dynamic _document;
        private readonly List<string> _files = new List<string>();
    }
}
