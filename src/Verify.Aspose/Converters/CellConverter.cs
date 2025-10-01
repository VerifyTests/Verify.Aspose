using Aspose.Cells;

class CellConverter :
    WriteOnlyJsonConverter<Cell>
{
    public override void Write(VerifyJsonWriter writer, Cell cell)
    {
        writer.WriteStartObject();

        writer.WriteMember(cell, cell.Name, "Name");
        writer.WriteMember(cell, cell.Row, "Row");
        writer.WriteMember(cell, cell.Column, "Column");
        writer.WriteMember(cell, cell.Value, "Value");
        writer.WriteMember(cell, cell.Formula, "Formula");
        writer.WriteMember(cell, cell.Type, "Type");

        writer.WriteEndObject();
    }
}