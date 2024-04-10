using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgeAssembly;
using SolidEdgePart;
using SolidEdgeGeometry;

namespace Remake_PMI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //Connect to the active instance
                SolidEdgeFramework.Application SolidEdgeClient = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                AssemblyDocument asm = SolidEdgeClient.ActiveDocument;
                PMI PMI = SolidEdgeClient.ActiveDocument.PMI;

                //Cycle through PMI dimensions
                foreach (Dimension dim in PMI.Dimensions)
                {
                    if (dim.StatusOfDimension == DimStatusConstants.seDimStatusError) continue; //Ignore broken PMI, cannot remake

                    dim.GetRelatedCount(out int relatedCount);

                    System.Data.DataTable related = new System.Data.DataTable();
                    related.Columns.Add("Object", typeof(object));
                    related.Columns.Add("X", typeof(double));
                    related.Columns.Add("Y", typeof(double));
                    related.Columns.Add("Z", typeof(double));
                    related.Columns.Add("KeyPoint", typeof(bool));
                    related.Columns.Add("RefObject", typeof(object));
                    related.Columns.Add("New", typeof(object));

                    //Store all attachment data
                    for (int i = 0; i < relatedCount; i++)
                    {
                        dim.GetRelated(i, out dynamic obj, out double X, out double Y, out double Z, out bool keyPoint);
                        object refObj = obj.Type == ((int)SolidEdgeFramework.ObjectType.igReference) ? obj.Object : obj;
                        related.Rows.Add((object)obj, X, Y, Z, keyPoint, refObj, null);
                    }

                    foreach (Occurrence occurrence in asm.Occurrences)
                    {
                        PartDocument part = occurrence.OccurrenceDocument;
                        Model model = part.Models.Item(1);
                        Body body = model.Body;

                        //Find the related edges through the part body
                        foreach (Edge edge in body.Edges[SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll]) //Method 1: Find by the edges collection
                        {
                            foreach (DataRow match in related.AsEnumerable().Where(x => x.Field<object>("RefObject") == edge))
                            {
                                object obj = match.Field<object>("Object"), refObj = match.Field<object>("RefObject");
                                if (obj == refObj)
                                {
                                    match.SetField(6, edge);
                                }
                                else
                                {
                                    object reference = asm.CreateReference(occurrence, edge);
                                    match.SetField(6, reference);
                                }
                            }
                            if (!related.AsEnumerable().Any(x => x.Field<object>("New") == null)) break;
                        }

                        if (related.AsEnumerable().Any(x => x.Field<object>("New") == null))
                        {
                            MessageBox.Show(dim.VariableTableName + ": Could not find edge(s)");
                            continue;
                        }

                         //Remake the dimension from the found edge
                        Dimensions dims = PMI.Dimensions;
                        DimInitData dimInitData = dims.DimInitData;

                        dimInitData.ClearParents();
                        dimInitData.SetType(dim.DimensionType);
                        dimInitData.SetAxisMode(dim.MeasurementAxisEx);
                        dimInitData.SetNumberOfParents(relatedCount);
                        dimInitData.SetPlane(dim.PMIPlane);

                        for (int i = 0; i < relatedCount; i++)
                        {
                            object obj = related.Rows[i].Field<object>("New");
                            double X = related.Rows[i].Field<double>("X");
                            double Y = related.Rows[i].Field<double>("Y");
                            double Z = related.Rows[i].Field<double>("Z");
                            bool keyPoint = related.Rows[i].Field<bool>("KeyPoint");
                            dimInitData.SetParentByIndex(i, obj, keyPoint, false, false, false, X, Y, Z);
                        }
                        try
                        {
                            Dimension dimNew = dims.AddDimension(dimInitData);
                            dimNew.VariableTableName = dim.VariableTableName + "_REMADE";
                        }
                        catch (Exception ex) 
                        {
                            MessageBox.Show(dim.VariableTableName + ": Could not remake");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error");
            }
        }
    }
}
