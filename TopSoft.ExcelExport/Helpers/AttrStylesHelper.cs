using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TopSoft.ExcelExport.Attributes;

namespace TopSoft.ExcelExport.Helpers
{
    static class AttrStylesHelper
    {
        static public Font GetFont(this CellTextAttribute cellTextAttribute)
        {
            Font retFont = new Font();

            if(cellTextAttribute.Bold)
            {
                retFont.Append(new Bold());
            }

            if(cellTextAttribute.Italic)
            {
                retFont.Append(new Italic());
            }

            if(cellTextAttribute.Underliine)
            {
                retFont.Append(new Underline());
            }
            return retFont;
        }

        static public PatternFill GetFill(this CellFillAttribute cellFillAttribute)
        {
             PatternFill retFill = new PatternFill() { PatternType = PatternValues.Solid };

            if(!string.IsNullOrEmpty(cellFillAttribute.HexColor))
            {
                
                retFill.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString(cellFillAttribute.HexColor) };
            }

            return retFill;
        }

        static public Border GetBorder(this CellBorderAttribute cellFontAttribute)
        {
            Border retBorder = new Border();

            if(cellFontAttribute.LeftBorder)
            {
                retBorder.LeftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
                retBorder.LeftBorder.Color = new Color() { Indexed = 64U };
            }

            if(cellFontAttribute.RightBorder)
            {
                retBorder.RightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
                retBorder.RightBorder.Color = new Color() { Indexed = 64U };
            }

            if(cellFontAttribute.TopBorder)
            {
                retBorder.TopBorder = new TopBorder() { Style = BorderStyleValues.Thin };
                retBorder.TopBorder.Color = new Color() { Indexed = 64U };
            }

            if(cellFontAttribute.BottomBorder)
            {
                retBorder.BottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
                retBorder.BottomBorder.Color = new Color() { Indexed = 64U };
            }

            if(cellFontAttribute.DiagonalBorder)
            {
                retBorder.DiagonalBorder = new DiagonalBorder() { Style = BorderStyleValues.Thin };
                retBorder.DiagonalBorder.Color = new Color() { Indexed = 64U };
            }

            return retBorder;
        }
    }
}
