using WordXMLDom;
using WordXMLDom.Exceptions;
namespace WordXMLDomTests.WordDocumentTests;

public class WordDocumentTests
{
    private readonly string pathToValidWordDoc = $@"{Directory.GetCurrentDirectory()}\..\..\..\WordTestFiles\TestDoc.docx";    
    private readonly string pathToInvalidWordDoc = $@"{Directory.GetCurrentDirectory()}\..\..\..\WordTestFiles\FakeDoc.txt";
    private readonly string pathToNonExistentFile = $@"{Directory.GetCurrentDirectory()}\..\..\..\WordTestFiles\NotFound.txt";
    private readonly string freshlyCreatedWordDoc = $@"{Directory.GetCurrentDirectory()}\..\..\..\WordTestFiles\FreshWordDoc.docx";


    // [Fact]
    // public void OpenDocument_WithValidPathAndValidWordDocument_ShouldCreateNewWordDocumentObj()
    // {
    //     // arrange
    //     WordDocument doc = new();

    //     // act
    //     doc.OpenDocument(pathToValidWordDoc);
    // }    

    [Fact]
    public void OpenDocument_FileDoesntExist_ShouldThrowFileNotFoundException()
    {
        // arrange
        WordDocument doc = new();

        // act
        void action() { doc.OpenDocument(pathToNonExistentFile); }


        // assert
        Assert.Throws<FileNotFoundException>(action);

    }

    [Fact]
    public void OpenDocument_WithInvalidFileType_ShouldThrowNotValidDocumentException()
    {
        // arrange
        WordDocument doc = new();

        // act
        void action() {doc.OpenDocument(pathToInvalidWordDoc);}

        // assert
        Assert.Throws<NotValidDocumentException>(action);
    }

    [Fact]
    public void OpenDocument_WithEmptyWordDocument_ShouldThrowNotValidDocumentException()
    {
        // arrange
        WordDocument doc = new();

        // act
        void action() {doc.OpenDocument(freshlyCreatedWordDoc);}

        // assert
        Assert.Throws<EmptyDocumentException>(action);
    }

    [Fact]
    public void printToXmlFile_WithValidDocumentPath()
    {
        // arrange
        WordDocument doc = new(pathToValidWordDoc);
        string path = $@"{Directory.GetCurrentDirectory()}\..\..\..\WordTestFiles\document.xml";
        if(File.Exists(path))
            File.Delete(path);

        // act
        doc.printToXmlFile(path);

        // assert
        Assert.True(File.Exists(path));        

    }
}