using Aspose.Cells;

class CellAreaConverter :
    WriteOnlyJsonConverter<CellArea>
{
    public override void Write(VerifyJsonWriter writer, CellArea area)
    {
        writer.WriteStartObject();

        if (area.StartColumn == area.EndColumn &&
            area.StartRow == area.EndRow)
        {
            writer.WriteMember(area, area.StartColumn, "Column");
            writer.WriteMember(area, area.StartRow, "Row");
        }
        else
        {
            writer.WriteMember(area, area.StartColumn, "StartColumn");
            writer.WriteMember(area, area.StartRow, "StartRow");
            writer.WriteMember(area, area.EndColumn, "EndColumn");
            writer.WriteMember(area, area.EndRow, "EndRow");
        }

        writer.WriteEndObject();
    }
}