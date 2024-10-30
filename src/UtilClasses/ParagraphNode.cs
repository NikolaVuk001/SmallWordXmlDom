namespace WordXMLDom.UtilClasses;
public class ParagraphNode(WordParagraph value)
{    
    public WordParagraph Value { get; set; } = value;
    public ParagraphNode? Previous { get; set; } = null;
    public ParagraphNode? Next { get; set; } = null;
}