using Aspose.Cells;

class HyperlinkConverter :
    WriteOnlyJsonConverter<Hyperlink>
{
    public override void Write(VerifyJsonWriter writer, Hyperlink link)
    {
        writer.WriteStartObject();

        writer.WriteMember(link, link.Address, "Address");
        // Aspose.Cells throws NullReferenceException from Hyperlink.TextToDisplay
        // when the hyperlink was created without display text (e.g. by OpenXml)
        try
        {
            writer.WriteMember(link, link.TextToDisplay, "Text");
        }
        catch (NullReferenceException)
        {
        }

        writer.WriteMember(link, link.Area, "Area");
        writer.WriteMember(link, link.LinkType, "LinkType");
        writer.WriteMember(link, link.ScreenTip, "ScreenTip");

        writer.WriteEndObject();
    }
}