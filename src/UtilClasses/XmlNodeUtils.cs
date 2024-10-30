using System.Text;
using System.Xml;
using WordXMLDom.Exceptions;

public static class XmlUtils
{

    public static string GetStringValueXml(XmlReader xmlReader)
    {
        Stack<string> elementsStack = new();
        StringBuilder sb = new();

        do
        {
            if(xmlReader.NodeType == XmlNodeType.EndElement)
            {
                
                if(elementsStack.Count > 0 && elementsStack.Peek() == xmlReader.Name)
                {
                    sb.Append($"""</{xmlReader.Name}>""");
                    elementsStack.Pop();
                    if(elementsStack.Count == 0)
                    {
                        break;
                    }
                }                
                else
                {
                    throw new NotValidDocumentException();
                }
                
            }

            if(xmlReader.NodeType == XmlNodeType.Element)
            {
                elementsStack.Push(xmlReader.Name);
                sb.Append($"""<{xmlReader.Name} """);

                // Write all the attributes;                         
                GetTagAttributes(xmlReader, sb);
                if(xmlReader.IsEmptyElement)
                {                
                    sb.Append("/>");
                    if(elementsStack.Count > 0 && elementsStack.Peek() == xmlReader.Name)
                    {                        
                        elementsStack.Pop();
                        if(elementsStack.Count == 0)
                        {
                            break;
                        }
                    }     
                    
                }
                else
                {
                    sb[sb.Length - 1] = '>';
                }                                
            }            
            xmlReader.Read();
        }while(elementsStack.Count != 0);
        
        return sb.ToString();
    }


    public static void GetTagAttributes(XmlReader xmlReader, StringBuilder sb)
    {
        var attrCout = xmlReader.AttributeCount;            
        if(attrCout > 0)
        {
            for(int i = 0; i < attrCout; i++)
            {
                xmlReader.MoveToAttribute(i);
                sb.Append($"""{xmlReader.Name}="{xmlReader.Value}" """);
            }
            xmlReader.MoveToElement();
        }
    }
    public static string GetTagAttributes(XmlReader xmlReader)
    {
        StringBuilder sb = new();
        var attrCout = xmlReader.AttributeCount;            
        if(attrCout > 0)
        {
            for(int i = 0; i < attrCout; i++)
            {
                xmlReader.MoveToAttribute(i);
                sb.Append($"""{xmlReader.Name}="{xmlReader.Value}" """);
            }
            xmlReader.MoveToElement();
        }

        return sb.ToString();
    }
}