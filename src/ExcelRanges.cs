using System;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        // Excel's Range[string] parser can reject fragmented unions even when
        // their A1 text fits the creation batch. Construct each local area and
        // combine actual Range objects instead of re-parsing a long union.
        object RangeAt(object sheet, string address)
        {
            if (address.IndexOf(',') < 0) return ((dynamic)sheet).Range[address];
            object combined = null;
            try
            {
                foreach (string area in address.Split(','))
                {
                    object part = null, previous = null;
                    try
                    {
                        if (String.IsNullOrWhiteSpace(area)) throw new ArgumentException("Zone de plage vide.", "address");
                        part = ((dynamic)sheet).Range[area];
                        if (combined == null) { combined = part; part = null; }
                        else
                        {
                            previous = combined; combined = null;
                            combined = ((dynamic)application).Union(previous, part);
                        }
                    }
                    finally { Release(previous); Release(part); }
                }
                object result = combined; combined = null; return result;
            }
            finally { Release(combined); }
        }
    }
}
