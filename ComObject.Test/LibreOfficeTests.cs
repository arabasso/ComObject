using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace ComObject.Test
{
    public class LibreOfficeTests
        : IDisposable
    {
        [Fact]
        public void Load_xtext_document()
        {
            var odt = GetTempFile(".odt");

            File.Copy(@"Files\Document.odt", odt);

            var xTextDocument = LoadXTextDocument(odt);

            var xParaCursor = xTextDocument.getText().createTextCursor();

            xParaCursor.gotoStartOfParagraph(false);
            xParaCursor.gotoEndOfParagraph(true);

            Assert.Equal(xParaCursor.getString(), "Paragraph 1");
        }

        [Fact]
        public void Save_xtext_document()
        {
            var odt = GetTempFile(".odt");

            var xTextDocument = NewXTextDocument();

            xTextDocument.storeAsURL( new Uri(odt).ToString(), new object[0]);

            Assert.True(File.Exists(odt));
        }

        [Theory]
        [InlineData(".docx", "MS Word 2007 XML")]
        [InlineData(".doc", "MS Word 97")]
        [InlineData(".pdf", "writer_pdf_Export")]
        public void Export_xtext_document(
            string extension,
            string filterName)
        {
            var docx = GetTempFile(extension);

            var xTextDocument = NewXTextDocument();

            var prop = _serviceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue");

            prop.Name = "FilterName";
            prop.Value = filterName;

            xTextDocument.storeToURL(new Uri(docx).ToString(), new [] { prop });

            Assert.True(File.Exists(docx));
        }

        [Fact]
        public void Add_paragraph()
        {
            const int PARAGRAPH_BREAK = 0;

            var xTextDocument = NewXTextDocument();

            var xText = xTextDocument.getText();

            const string text = "Pagragraph 1";

            xText.insertString(xText.getEnd(), text, false);
            xText.insertControlCharacter(xText.getEnd(), PARAGRAPH_BREAK, false);

            var xParaCursor = xTextDocument.getText().createTextCursor();

            xParaCursor.gotoStartOfParagraph(false);
            xParaCursor.gotoEndOfParagraph(true);

            Assert.Equal(xParaCursor.getString(), text);
        }

        [Theory]
        [InlineData("Text", "TextReplace")]
        public void Replace_text(
            string search,
            string replace)
        {
            const int PARAGRAPH_BREAK = 0;

            var xTextDocument = NewXTextDocument();

            var xText = xTextDocument.getText();

            xText.insertString(xText.getEnd(), search, false);
            xText.insertControlCharacter(xText.getEnd(), PARAGRAPH_BREAK, false);

            var xReplaceDescr = xTextDocument.createReplaceDescriptor();

            xReplaceDescr.setSearchString(search);
            xReplaceDescr.setReplaceString(replace);

            xTextDocument.replaceAll(xReplaceDescr);

            var xParaCursor = xTextDocument.getText().createTextCursor();

            xParaCursor.gotoStartOfParagraph(false);
            xParaCursor.gotoEndOfParagraph(true);

            Assert.Equal(xParaCursor.getString(), replace);
        }

        private dynamic NewXTextDocument()
        {
            var xTextDocument = _desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, _loadProps);

            _documents.Add(xTextDocument);

            return xTextDocument;
        }

        private dynamic LoadXTextDocument(
            string path)
        {
            var xTextDocument = _desktop.loadComponentFromURL(new Uri(path).ToString(), "_blank", 0, _loadProps);

            _documents.Add(xTextDocument);

            return xTextDocument;
        }

        private string GetTempFile(
            string extension)
        {
            var file = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetTempFileName(), extension));

            _files.Add(file);

            return file;
        }

        public LibreOfficeTests()
        {
            _serviceManager = new ComObject("com.sun.star.ServiceManager");

            _desktop = _serviceManager.createInstance("com.sun.star.frame.Desktop");

            var prop1 = _serviceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue");

            prop1.Name = "Hidden";
            prop1.Value = true;

            var prop2 = _serviceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue");

            prop2.Name = "ShowTrackedChanges";
            prop2.Value = false;

            _loadProps = new object[]
            {
                prop1,
                prop2
            };
        }

        public void Dispose()
        {
            foreach (var document in _documents)
            {
                document.close(false);
                document.Dispose();
            }

            _desktop.Dispose();

            _serviceManager.Dispose();

            foreach (var file in _files.Where(File.Exists))
            {
                File.Delete(file);
            }
        }

        private readonly dynamic _serviceManager;
        private readonly dynamic _desktop;
        private readonly object [] _loadProps;
        private readonly List<string> _files = new List<string>();
        private readonly List<dynamic> _documents = new List<dynamic>();
    }
}
