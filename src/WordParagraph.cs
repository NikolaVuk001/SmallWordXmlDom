using System.Text;
using System.Xml;

namespace WordXMLDom;

// Tag <w:r></w:r>
public class WordParagraph()
{
    public string Attributes = "";
    public string Props { get; set; } = "";
    public string MiscElemens { get; set; } = "";    
    public List<TextRun> TextRuns { get; set; } = new();

    public string FirstWord { get; set; } = "";

    public static WordParagraph CreateParagraphFromXml(XmlReader xmlReader)
    {        
        WordParagraph wordParagraph = new();
        var paraID = xmlReader.GetAttribute(0);
        if(xmlReader.IsEmptyElement)
        {
            return wordParagraph;
        }        
        wordParagraph.Attributes = XmlUtils.GetTagAttributes(xmlReader);
        
        do
        {
            xmlReader.Read();
            if(xmlReader.Name == "w:pPr" && xmlReader.NodeType == XmlNodeType.Element)
            {
                // StringBuilder sb = new();                        
                // XmlUtils.GetStringValueXml(xmlReader, "w:pPr", sb);
                wordParagraph.Props = XmlUtils.GetStringValueXml(xmlReader);
                // sb.ToString();
                continue;
                
            }

             // Text Runs
            if(xmlReader.Name == "w:r" && xmlReader.NodeType == XmlNodeType.Element)
            {
                TextRun textRun = new();                
                wordParagraph.AddTextRunsFromXml(xmlReader);
                continue;                                                      
            }     

            // Elements before text run
            if( xmlReader.Name != "w:r" &&
                xmlReader.Name != "w:pPr" &&
                xmlReader.NodeType == XmlNodeType.Element)
              {                
                wordParagraph.MiscElemens = XmlUtils.GetStringValueXml(xmlReader);             
                continue;
                
              }

            

        } while(xmlReader.NodeType != XmlNodeType.EndElement || xmlReader.Name != "w:p");

        return wordParagraph;        
    }

    private void AddTextRunsFromXml(XmlReader xmlReader)
    {
        TextRun textRun = new();
        textRun.Attributes = XmlUtils.GetTagAttributes(xmlReader);
        do 
        {
            xmlReader.Read();

            // Props
            if(xmlReader.Name == "w:rPr")
            {
                StringBuilder sb = new();
                textRun.Props = XmlUtils.GetStringValueXml(xmlReader);
                 
                continue;
            }                            

            // Text Piece
            if(xmlReader.Name == "w:t" && xmlReader.NodeType == XmlNodeType.Element)
            {
                textRun.AddTextPieceFromXml(xmlReader);
                continue;
                
            }                            
            
        }        
        while(xmlReader.NodeType != XmlNodeType.EndElement || xmlReader.Name != "w:t");
        
        this.TextRuns.Add(textRun);
        // getting the first word of first text Piece
        
        if(textRun.TextPieces.Count > 0 && FirstWord == "")
        {
            FirstWord = StringUtils.GetFirstWord(textRun.TextPieces[0].Text);
        }
        
        
    }

    
   
}



// tag <w:r></w:r>
public class TextRun
{

    public string Attributes { get; set; } = "";
    public string Props { get; set; } = "";
    public string TextTag { get; set; } = "<w:t></w:t>";
    public List<TextPiece> TextPieces { get; set; } = new();    


    public void AddTextPieceFromXml(XmlReader xmlReader)
    {
        TextPiece textPiece = new();        
        textPiece.Attributes = XmlUtils.GetTagAttributes(xmlReader);        
        // Move to text with Read
        xmlReader.Read();
        textPiece.Text = xmlReader.Value;

        this.TextPieces.Add(textPiece);
        
    }
}
// tag <w:t></w:t>
public class TextPiece
{
    public string Attributes { get; set; } = "";
    public string Text { get; set; } = "";
}



