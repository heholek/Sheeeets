using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VVVV.PluginInterfaces.V2;
using VVVV.Utils.VMath;

namespace Sheeeets.Nodes
{
    [PluginInfo(
        Name = "GetCellValue",
        Category = "Sheeeets",
        Author = "microdee"
    )]
    public class SheeeetsGetCellValueNode : IPluginEvaluate
    {
        [Input("Sheet")] public Pin<object> FSheet;
        [Input("Row Col ")] public ISpread<ISpread<Vector2D>> FRowCol;

        [Output("Value")] public ISpread<ISpread<string>> FVal;
        [Output("Note")] public ISpread<ISpread<string>> FNote;

        public void Evaluate(int SpreadMax)
        {
            if (FSheet.IsConnected && (FSheet.SliceCount > 0))
            {
                FVal.SliceCount = FSheet.SliceCount;
                FNote.SliceCount = FSheet.SliceCount;
                for (int i = 0; i < FVal.SliceCount; i++)
                {
                    if (FSheet[i] is Sheet)
                    {
                        var sheet = (Sheet)FSheet[i];
                        var sm = FRowCol[i].SliceCount;
                        FVal[i].SliceCount = sm;
                        FNote[i].SliceCount = sm;
                        for (int j = 0; j < sm; j++)
                        {
                            var cell = sheet[(int)FRowCol[i][j].x, (int)FRowCol[i][j].y];
                            if (cell == null) continue;
                            FVal[i][j] = cell.FormattedValue;
                            FNote[i][j] = cell.Note;
                        }
                    }
                    else
                    {
                        FVal[i].SliceCount = 0;
                        FNote[i].SliceCount = 0;
                    }
                }
            }
            else
            {
                FVal.SliceCount = 0;
                FNote.SliceCount = 0;
            }
        }
    }
    [PluginInfo(
        Name = "OccupiedGridDimensions",
        Category = "Sheeeets",
        Author = "microdee"
    )]
    public class SheeeetsOccupiedGridDimensionsNode : IPluginEvaluate
    {
        [Input("Sheet")]
        public Pin<object> FSheet;

        [Output("Dimension")]
        public ISpread<Vector2D> FRowCols;

        public void Evaluate(int SpreadMax)
        {
            if (FSheet.IsConnected && (FSheet.SliceCount > 0))
            {
                FRowCols.SliceCount = FSheet.SliceCount;
                for (int i = 0; i < FSheet.SliceCount; i++)
                {
                    if (!(FSheet[i] is Sheet)) continue;
                    var sheet = (Sheet)FSheet[i];
                    FRowCols[i] = new Vector2D(sheet.OccupiedGridHeight, sheet.OccupiedGridWidth);
                }
            }
            else
            {
                FRowCols.SliceCount = 0;
            }
        }
    }
    [PluginInfo(
        Name = "MergedRanges",
        Category = "Sheeeets",
        Author = "microdee"
    )]
    public class SheeeetsMergedRangesNode : IPluginEvaluate
    {
        [Input("Sheet")]
        public Pin<object> FSheet;

        [Output("Start")]
        public ISpread<ISpread<Vector2D>> FTL;
        [Output("End")]
        public ISpread<ISpread<Vector2D>> FBR;

        public void Evaluate(int SpreadMax)
        {
            if (FSheet.IsConnected && (FSheet.SliceCount > 0))
            {
                FTL.SliceCount = FSheet.SliceCount;
                FBR.SliceCount = FSheet.SliceCount;
                for (int i = 0; i < FSheet.SliceCount; i++)
                {
                    if (FSheet[i] is Sheet)
                    {
                        var sheet = (Sheet)FSheet[i];

                        if (sheet.GoogleSheet.Merges != null)
                        {
                            var mc = sheet.GoogleSheet.Merges.Count;
                            FTL[i].SliceCount = mc;
                            FBR[i].SliceCount = mc;
                            for (int j = 0; j < mc; j++)
                            {
                                int tlr = -1, tlc = -1, brr = -1, brc = -1;
                                var merge = sheet.GoogleSheet.Merges[j];
                                if (merge.StartRowIndex.HasValue)
                                {
                                    tlr = merge.StartRowIndex.Value;
                                    tlc = merge.StartColumnIndex.Value;
                                }
                                if (merge.EndRowIndex.HasValue)
                                {
                                    brr = merge.EndRowIndex.Value;
                                    brc = merge.EndColumnIndex.Value;
                                }
                                FTL[i][j] = new Vector2D(tlr, tlc);
                                FBR[i][j] = new Vector2D(brr, brc);
                            }
                        }
                        else
                        {
                            FTL[i].SliceCount = 0;
                            FBR[i].SliceCount = 0;
                        }
                    }
                    else
                    {
                        FTL[i].SliceCount = 0;
                        FBR[i].SliceCount = 0;
                    }
                }
            }
            else
            {
                FTL.SliceCount = 0;
                FBR.SliceCount = 0;
            }
        }
    }
}
