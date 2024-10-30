using System.IO.Compression;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using Microsoft.VisualBasic;
using WordXMLDom.Consts;
using WordXMLDom.Exceptions;
using WordXMLDom.UtilClasses;

namespace WordXMLDom;
public class WordDocument 
{
    public int ParagraphCount { get; set; } = 0;

    private string DocumentAttrs {get; set;} = "";

    private string Sections { get; set; } = "";
    private ParagraphNode? head = null;
    private ParagraphNode? middle = null;
    private ParagraphNode? tail = null;
    private Dictionary<string, List<int>> firstWordIndexes = new();

    /// <summary>
    /// Initializes empty WordDocument object.
    /// </summary>    
    public WordDocument()
    {
        // Add intitialization For all fields        
    }    

    /// <summary>
    /// Initializes WordDocument object with all the attributes from the given file.
    /// </summary>
    /// <exception cref="EmptyDocumentException">The contents of the .docx file couldn't be interpreted as a Zip archive.</exception>
    /// <exception cref="FileNotFoundException">File couldn't be found at the specifed path.</exception>
    /// <exception cref="NotValidDocumentException">File that is provided doesnt have .docx extension</exception>
    /// <exception cref="ArgumentException">The stream is already closed or does not support reading.</exception>
    /// <exception cref="ArgumentNullException">The stream is null.</exception>    
    /// <param name="path">The path to the Word(.docx) document.</param>    
    public WordDocument(string path)
    {
       this.OpenDocument(path);        
    }    

    /// <summary>
    /// Loads the .docx file to a WordDocument object.
    /// </summary>
    /// <exception cref="EmptyDocumentException">The contents of the .docx file couldn't be interpreted as a Zip archive.</exception>
    /// <exception cref="FileNotFoundException">File couldn't be found at the specifed path.</exception>
    /// <exception cref="NotValidDocumentException">File that is provided doesnt have .docx extension</exception>
    /// <exception cref="ArgumentException">The stream is already closed or does not support reading.</exception>
    /// <exception cref="ArgumentNullException">The stream is null.</exception>    
    /// <param name="path">The path to the Word(.docx) document.</param>    
    public void OpenDocument(string path)
    {
         if(!File.Exists(path))
        {
            throw new FileNotFoundException();
        }
        if(path[^5..] != ".docx")
        {
            throw new NotValidDocumentException();
        }
        using FileStream stream = new(path: path, FileMode.Open);

        ZipArchive zip;
        try
        {
            zip = new(stream);
        }
        catch(InvalidDataException)
        {
            throw new EmptyDocumentException();
        }        

        var documentStream = zip.Entries[2].Open(); // Check if the document.xml is always at index of 2?
        XmlReaderSettings settings = new()
        {
            IgnoreWhitespace = false
        };
        using XmlReader xmlReader = XmlReader.Create(documentStream, settings);

        
        CheckIfValidWordDocument(xmlReader);
        var paraCount = 0;
        while(xmlReader.Read())
        {
            // assumption that the word document has w:p as the first element after the body
            
            if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "w:p")
            {                            
                ParagraphNode node = CreateParagraphNode();                
                node.Value = WordParagraph.CreateParagraphFromXml(xmlReader);                
                paraCount++;
                // append paragraph
                Append(node);
            }    
            if(xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "w:sectPr")
            {
                Sections = XmlUtils.GetStringValueXml(xmlReader);
            }   
        }                                      
    }

    public void printToXmlFile(string path)
    {
        
        using StreamWriter w = new(path);

        w.WriteLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        w.WriteLine($"""<w:document {this.DocumentAttrs}>""");
        w.WriteLine("<w:body>");
        if(this.ParagraphCount != 0)
        {
            var curr = head;
            while(curr != null)
            {
                var paragraph = curr.Value;
                w.WriteLine($"""<w:p {paragraph.Attributes}>""");
                w.WriteLine($"{paragraph.Props}");
                w.WriteLine($"{paragraph.MiscElemens}");
                foreach(var textRun in paragraph.TextRuns)
                {                
                    w.WriteLine($"<w:r {textRun.Attributes}>");
                    w.WriteLine($"{textRun.Props}");
                    foreach(var textPiece in textRun.TextPieces)
                    {
                        w.WriteLine($"<w:t {textPiece.Attributes}>{textPiece.Text}</w:t>");
                    }
                    w.WriteLine("</w:r>");
                }
                w.WriteLine("</w:p>");
                curr = curr.Next;
            }
        }
        w.WriteLine($"{Sections}");
        w.WriteLine("</w:body>");
        w.WriteLine("</w:document>");

        w.Close();
    }

    private void Append(ParagraphNode node)
    {    
        if(ParagraphCount++ == 0)
        {
            head = tail = middle = node;
            head.Next = middle;
            middle.Previous = head;
            middle.Next = tail;
            tail.Previous = middle;
            
        }
        else {
            node.Previous = tail;
            tail!.Next = node;
            tail = node;
        }            

        if(ParagraphCount % 2 != 0)
        {
            middle = middle!.Next;
        }     
        if(firstWordIndexes.ContainsKey(node.Value.FirstWord))
        {
            firstWordIndexes[node.Value.FirstWord].Add(ParagraphCount - 1);
        }
        else
        {
            firstWordIndexes.Add(node.Value.FirstWord, new() {ParagraphCount - 1});
        }
        
          
    }

    private void CheckIfValidWordDocument(XmlReader xmlReader)
    {
        xmlReader.Read();
        if(xmlReader.NodeType != XmlNodeType.XmlDeclaration && xmlReader.Value.ToString() != TagConstants.XMLTagValue)
        {
            throw new NotValidDocumentException();
        }
        xmlReader.Read();
        if(xmlReader.NodeType == XmlNodeType.Whitespace)
        {
            xmlReader.Read();
        }        
        
        if(xmlReader.Name != "w:document")
        {
            throw new NotValidDocumentException("Document doesnt have valid structure(w:document tag not found)");
        }
        this.DocumentAttrs = XmlUtils.GetTagAttributes(xmlReader);
        xmlReader.Read();
        if(xmlReader.Name != "w:body")
        {
            throw new NotValidDocumentException("Document doesnt have valid structure(w:body tag not found)");
        }                        
    }

    private ParagraphNode CreateParagraphNode()
    {
        return new(new());                
    }
    

    
    
}