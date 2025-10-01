using Aspose.Cells;

class HyperlinkConverter :
    WriteOnlyJsonConverter<Hyperlink>
{
    public override void Write(VerifyJsonWriter writer, Hyperlink link)
    {
        writer.WriteStartObject();

        writer.WriteMember(link, link.Address, "Address");
        writer.WriteMember(link, link.TextToDisplay, "Text");
        writer.WriteMember(link, link.Area, "Area");
        writer.WriteMember(link, link.LinkType, "LinkType");
        writer.WriteMember(link, link.ScreenTip, "ScreenTip");

        writer.WriteEndObject();
    }
}