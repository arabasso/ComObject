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
        public void Open_document()
        {
            var docx = GetTempFile(".docx");

            File.Copy(@"Files\Document.docx", docx);

            var document = OpenDocument(docx);

            Assert.Equal("Paragraph 1", document.Paragraphs[1].Range.Text.Trim());
        }

        [Fact]
        public void Save_document()
        {
            var docx = GetTempFile(".docx");

            var document = NewDocument();

            document.SaveAs(docx);

            Assert.True(File.Exists(docx));
        }

        [Fact]
        public void Save_document_as_doc()
        {
            const int wdFormatDocument = 0;

            var doc = GetTempFile(".doc");

            var document = NewDocument();

            document.SaveAs(doc, wdFormatDocument);

            Assert.True(File.Exists(doc));
        }

        [Fact]
        public void Save_document_as_pdf()
        {
            const int wdFormatPDF = 17;

            var pdf = GetTempFile(".pdf");

            var document = NewDocument();

            document.SaveAs(pdf, wdFormatPDF);

            Assert.True(File.Exists(pdf));
        }

        [Fact]
        public void Add_paragraph()
        {
            var document = NewDocument();

            var paragraph = document.Paragraphs.Add();

            const string text = "Pagragraph 1";

            paragraph.Range.Text = text;

            Assert.Equal(document.Paragraphs[1].Range.Text.Trim(), text);
        }

        [Theory]
        [InlineData("Text", "TextReplace")]
        public void Replace_text(
            string search,
            string replace)
        {
            const int wdFindContinue = 1;
            const int wdReplaceAll = 2;

            var document = NewDocument();

            var paragraph = document.Paragraphs.Add();

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

            Assert.Equal(replace, document.Paragraphs[1].Range.Text.Trim());
        }

        public WordInteropTests()
        {
            _application = new ComObject("Word.Application");

            _application.Visible = false;
        }

        private string GetTempFile(
            string extension)
        {
            var file = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetTempFileName(), extension));

            _files.Add(file);

            return file;
        }

        private dynamic NewDocument()
        {
            var document = _application.Documents.Add();

            _documents.Add(document);

            return document;
        }

        private dynamic OpenDocument(
            string path)
        {
            var document = _application.Documents.Open(Path.GetFullPath(path));

            _documents.Add(document);

            return document;
        }

        public void Dispose()
        {
            foreach (var document in _documents)
            {
                document.Close(false);
                document.Dispose();
            }

            _application.Quit(false);
            _application.Dispose();

            foreach (var file in _files.Where(File.Exists))
            {
                File.Delete(file);
            }
        }

        private readonly dynamic _application;
        private readonly List<string> _files = new List<string>();
        private readonly List<dynamic> _documents = new List<dynamic>();
    }
}
