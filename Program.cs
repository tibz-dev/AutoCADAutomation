using System;
using System.IO;
using System.Threading;

namespace AutoCADAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Select Mode:");
                Console.WriteLine("1 - Electrical");
                Console.WriteLine("2 - Civil");
                int mode = Convert.ToInt32(Console.ReadLine());

                Console.Write("Enter room width (mm): ");
                double width = Convert.ToDouble(Console.ReadLine());

                Console.Write("Enter room height (mm): ");
                double height = Convert.ToDouble(Console.ReadLine());

                int lights = 0;
                int switches = 0;

                if (mode == 1)
                {
                    Console.Write("Enter number of lights: ");
                    lights = Convert.ToInt32(Console.ReadLine());

                    Console.Write("Enter number of switches: ");
                    switches = Convert.ToInt32(Console.ReadLine());
                }

                string folder = @"C:\Users\sesha\AutoCADProjects";

                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string outputPath = $@"{folder}\Project_{DateTime.Now.Ticks}.dwg";

                // Start AutoCAD
                Type acadType = Type.GetTypeFromProgID("AutoCAD.Application");
                dynamic acad = Activator.CreateInstance(acadType);
                acad.Visible = false;

                Thread.Sleep(5000);

                dynamic doc = acad.Documents.Add();
                dynamic modelSpace = doc.ModelSpace;

                Thread.Sleep(2000);

                // ===== DRAW ROOM =====
                double[] p1 = { 0, 0, 0 };
                double[] p2 = { width, 0, 0 };
                double[] p3 = { width, height, 0 };
                double[] p4 = { 0, height, 0 };

                modelSpace.AddLine(p1, p2);
                modelSpace.AddLine(p2, p3);
                modelSpace.AddLine(p3, p4);
                modelSpace.AddLine(p4, p1);

                if (mode == 1)
                {
                    // ================= ELECTRICAL =================

                    // ===== DB BOARD =====
                    double[] dbPos = { 200, height - 300, 0 };
                    modelSpace.AddLine(
                        new double[] { dbPos[0], dbPos[1], 0 },
                        new double[] { dbPos[0] + 300, dbPos[1], 0 });
                    modelSpace.AddLine(
                        new double[] { dbPos[0] + 300, dbPos[1], 0 },
                        new double[] { dbPos[0] + 300, dbPos[1] - 300, 0 });
                    modelSpace.AddLine(
                        new double[] { dbPos[0] + 300, dbPos[1] - 300, 0 },
                        new double[] { dbPos[0], dbPos[1] - 300, 0 });
                    modelSpace.AddLine(
                        new double[] { dbPos[0], dbPos[1] - 300, 0 },
                        new double[] { dbPos[0], dbPos[1], 0 });

                    modelSpace.AddText("DB", new double[] { dbPos[0], dbPos[1] + 100, 0 }, 150);

                    // ===== LIGHTS =====
                    double spacing = width / (lights + 1);
                    double[][] lightPositions = new double[lights][];

                    for (int i = 0; i < lights; i++)
                    {
                        double x = spacing * (i + 1);
                        double y = height / 2;

                        double[] pos = { x, y, 0 };
                        lightPositions[i] = pos;

                        // circle
                        modelSpace.AddCircle(pos, 200);

                        // cross (symbol)
                        modelSpace.AddLine(
                            new double[] { x - 150, y - 150, 0 },
                            new double[] { x + 150, y + 150, 0 });
                        modelSpace.AddLine(
                            new double[] { x - 150, y + 150, 0 },
                            new double[] { x + 150, y - 150, 0 });

                        modelSpace.AddText($"L{i + 1}", new double[] { x + 100, y + 100, 0 }, 150);
                    }

                    // ===== SWITCHES =====
                    double[][] switchPositions = new double[switches][];

                    for (int i = 0; i < switches; i++)
                    {
                        double x = 200;
                        double y = (height / (switches + 1)) * (i + 1);

                        double[] pos = { x, y, 0 };
                        switchPositions[i] = pos;

                        // rectangle switch
                        modelSpace.AddLine(
                            new double[] { x - 100, y - 100, 0 },
                            new double[] { x + 100, y - 100, 0 });
                        modelSpace.AddLine(
                            new double[] { x + 100, y - 100, 0 },
                            new double[] { x + 100, y + 100, 0 });
                        modelSpace.AddLine(
                            new double[] { x + 100, y + 100, 0 },
                            new double[] { x - 100, y + 100, 0 });
                        modelSpace.AddLine(
                            new double[] { x - 100, y + 100, 0 },
                            new double[] { x - 100, y - 100, 0 });

                        modelSpace.AddText($"S{i + 1}", new double[] { x + 120, y + 120, 0 }, 150);

                        // main supply line from DB
                        modelSpace.AddLine(dbPos, pos);
                    }

                    // ===== SMART WIRING =====
                    int lightsPerSwitch = lights / switches;
                    int extraLights = lights % switches;
                    int lightIndex = 0;

                    for (int i = 0; i < switches; i++)
                    {
                        int assigned = lightsPerSwitch + (i < extraLights ? 1 : 0);

                        for (int j = 0; j < assigned; j++)
                        {
                            if (lightIndex < lights)
                            {
                                double[] sw = switchPositions[i];
                                double[] lt = lightPositions[lightIndex];

                                double[] mid = { sw[0], lt[1], 0 };

                                modelSpace.AddLine(sw, mid);
                                modelSpace.AddLine(mid, lt);

                                lightIndex++;
                            }
                        }
                    }
                }
                else
                {
                    // ================= CIVIL =================

                    // ===== DOOR (BOTTOM WALL OPENING) =====
                    double doorWidth = 800;

                    // Door gap (remove part of wall by drawing split lines)
                    modelSpace.AddLine(
                        new double[] { 0, 0, 0 },
                        new double[] { width / 2 - doorWidth / 2, 0, 0 });

                    modelSpace.AddLine(
                        new double[] { width / 2 + doorWidth / 2, 0, 0 },
                        new double[] { width, 0, 0 });

                    // Door swing arc
                    double[] doorCenter = new double[] { width / 2 - doorWidth / 2, 0, 0 };

                    modelSpace.AddArc(
                        doorCenter,
                        doorWidth,
                        0,
                        Math.PI / 2 // 90 degrees
                    );

                    // ===== WINDOW (TOP WALL OPENING) =====
                    double windowWidth = 1000;

                    // Left wall to window
                    modelSpace.AddLine(
                        new double[] { 0, height, 0 },
                        new double[] { width / 4, height, 0 });

                    // Right wall from window
                    modelSpace.AddLine(
                        new double[] { width / 4 + windowWidth, height, 0 },
                        new double[] { width, height, 0 });

                    // Window frame (double line)
                    modelSpace.AddLine(
                        new double[] { width / 4, height, 0 },
                        new double[] { width / 4 + windowWidth, height, 0 });

                    modelSpace.AddLine(
                        new double[] { width / 4, height - 100, 0 },
                        new double[] { width / 4 + windowWidth, height - 100, 0 });

                    // ===== LABELS =====
                    modelSpace.AddText("Door", new double[] { width / 2, -300, 0 }, 150);
                    modelSpace.AddText("Window", new double[] { width / 4, height + 200, 0 }, 150);
                }

                // ===== SAVE =====
                doc.SaveAs(outputPath);
                doc.Close(false);
                acad.Quit();

                Console.WriteLine("\nProject created successfully!");
                Console.WriteLine("Saved at: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}