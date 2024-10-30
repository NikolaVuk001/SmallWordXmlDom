public static class StringUtils
{
    public static string GetFirstWord(string input)
    {
        int firstSpaceIndex = input.IndexOf(' ');
        if(firstSpaceIndex == -1)
        {
            return input;
        }
        return input[..firstSpaceIndex];
    }
}