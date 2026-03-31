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
                Console.Write("Enter choice: ");
                int mode = Convert.ToInt32(Console.ReadLine());

                string folder = @"C:\Users\sesha\AutoCADProjects";
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string outputPath = $@"{folder}\Project_{DateTime.Now.Ticks}.dwg";

                Type acadType = Type.GetTypeFromProgID("AutoCAD.Application");
                if (acadType == null)
                {
                    Console.WriteLine("AutoCAD is not installed or COM ProgID was not found.");
                    return;
                }

                dynamic acad = Activator.CreateInstance(acadType);
                acad.Visible = false;

                Thread.Sleep(5000);

                dynamic doc = acad.Documents.Add();
                dynamic modelSpace = doc.ModelSpace;

                Thread.Sleep(2000);

                if (mode == 1)
                {
                    RunElectrical(modelSpace);
                }
                else if (mode == 2)
                {
                    RunCivil(modelSpace);
                }
                else
                {
                    Console.WriteLine("Invalid option.");
                    doc.Close(false);
                    acad.Quit();
                    return;
                }

                bool saved = false;
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        doc.SaveAs(outputPath);
                        saved = true;
                        break;
                    }
                    catch
                    {
                        Console.WriteLine("Retrying save...");
                        Thread.Sleep(3000);
                    }
                }

                doc.Close(false);
                acad.Quit();

                if (saved)
                {
                    Console.WriteLine("\nProject created successfully!");
                    Console.WriteLine("Saved at: " + outputPath);
                }
                else
                {
                    Console.WriteLine("❌ Failed to save drawing.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Error: " + ex.Message);
            }
        }

        static void RunElectrical(dynamic modelSpace)
        {
            Console.Write("Enter room width (mm): ");
            double width = Convert.ToDouble(Console.ReadLine());

            Console.Write("Enter room height (mm): ");
            double height = Convert.ToDouble(Console.ReadLine());

            Console.Write("Enter number of lights: ");
            int lights = Convert.ToInt32(Console.ReadLine());

            Console.Write("Enter number of switches: ");
            int switches = Convert.ToInt32(Console.ReadLine());

            double[] p1 = { 0, 0, 0 };
            double[] p2 = { width, 0, 0 };
            double[] p3 = { width, height, 0 };
            double[] p4 = { 0, height, 0 };

            modelSpace.AddLine(p1, p2);
            modelSpace.AddLine(p2, p3);
            modelSpace.AddLine(p3, p4);
            modelSpace.AddLine(p4, p1);

            double[] dbPos = { 200, height - 300, 0 };
            DrawRectangle(modelSpace, dbPos[0], dbPos[1] - 300, 300, 300);
            modelSpace.AddText("DB", new double[] { dbPos[0], dbPos[1] + 100, 0 }, 150);

            double spacing = width / (lights + 1);
            double[][] lightPositions = new double[lights][];

            for (int i = 0; i < lights; i++)
            {
                double x = spacing * (i + 1);
                double y = height / 2;
                double[] pos = { x, y, 0 };
                lightPositions[i] = pos;

                modelSpace.AddCircle(pos, 200);
                modelSpace.AddLine(
                    new double[] { x - 150, y - 150, 0 },
                    new double[] { x + 150, y + 150, 0 });
                modelSpace.AddLine(
                    new double[] { x - 150, y + 150, 0 },
                    new double[] { x + 150, y - 150, 0 });

                modelSpace.AddText($"L{i + 1}", new double[] { x + 120, y + 120, 0 }, 150);
            }

            double[][] switchPositions = new double[switches][];

            for (int i = 0; i < switches; i++)
            {
                double x = 200;
                double y = (height / (switches + 1)) * (i + 1);
                double[] pos = { x, y, 0 };
                switchPositions[i] = pos;

                DrawRectangle(modelSpace, x - 100, y - 100, 200, 200);
                modelSpace.AddText($"S{i + 1}", new double[] { x + 120, y + 120, 0 }, 150);

                modelSpace.AddLine(dbPos, pos);
            }

            int lightsPerSwitch = switches > 0 ? lights / switches : 0;
            int extraLights = switches > 0 ? lights % switches : 0;
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

        static void RunCivil(dynamic modelSpace)
        {
            Console.Write("Enter overall plan width (mm): ");
            double totalWidth = Convert.ToDouble(Console.ReadLine());

            Console.Write("Enter overall plan height (mm): ");
            double totalHeight = Convert.ToDouble(Console.ReadLine());

            Console.Write("Enter wall thickness (mm): ");
            double wallThickness = Convert.ToDouble(Console.ReadLine());

            Console.Write("Enter number of rooms: ");
            int roomCount = Convert.ToInt32(Console.ReadLine());

            if (roomCount < 1)
            {
                Console.WriteLine("Room count must be at least 1.");
                return;
            }

            DrawDoubleWallRoom(modelSpace, 0, 0, totalWidth, totalHeight, wallThickness);

            double roomWidth = totalWidth / roomCount;

            for (int i = 1; i < roomCount; i++)
            {
                double x = roomWidth * i;
                modelSpace.AddLine(
                    new double[] { x, wallThickness, 0 },
                    new double[] { x, totalHeight - wallThickness, 0 });
            }

            for (int i = 0; i < roomCount; i++)
            {
                double roomX = i * roomWidth;
                double innerX = roomX + wallThickness;
                double innerY = wallThickness;
                double innerWidth = roomWidth - (2 * wallThickness);
                double innerHeight = totalHeight - (2 * wallThickness);

                modelSpace.AddText(
                    $"Room {i + 1}",
                    new double[] { innerX + 300, totalHeight - 500, 0 },
                    180);

                Console.WriteLine($"\n--- Room {i + 1} Door Setup ---");
                Console.Write("Choose door wall (top/bottom/left/right): ");
                string doorWall = (Console.ReadLine() ?? "").Trim().ToLower();

                Console.Write("Door offset from wall start (mm): ");
                double doorOffset = Convert.ToDouble(Console.ReadLine());

                Console.Write("Door width (mm): ");
                double doorWidth = Convert.ToDouble(Console.ReadLine());

                DrawDoorForRoom(
                    modelSpace,
                    roomX,
                    0,
                    roomWidth,
                    totalHeight,
                    wallThickness,
                    doorWall,
                    doorOffset,
                    doorWidth
                );

                double windowWidth = 1000;
                DrawWindowForRoom(
                    modelSpace,
                    roomX,
                    0,
                    roomWidth,
                    totalHeight,
                    wallThickness,
                    "top",
                    Math.Max(300, (roomWidth - windowWidth) / 2),
                    windowWidth
                );

                Console.WriteLine($"\n--- Room {i + 1} Furniture ---");
                Console.Write("Add bed? (yes/no): ");
                string addBed = (Console.ReadLine() ?? "").Trim().ToLower();

                if (addBed == "yes")
                {
                    DrawBed(modelSpace, innerX + 300, innerY + 300, 2000, 1500);
                }

                Console.Write("Add table? (yes/no): ");
                string addTable = (Console.ReadLine() ?? "").Trim().ToLower();

                if (addTable == "yes")
                {
                    DrawTable(modelSpace, innerX + innerWidth - 1200, innerY + 400, 800, 800);
                }

                Console.Write("Add toilet set? (yes/no): ");
                string addToilet = (Console.ReadLine() ?? "").Trim().ToLower();

                if (addToilet == "yes")
                {
                    DrawToilet(modelSpace, innerX + innerWidth - 900, innerY + innerHeight - 1400);
                }
            }
        }

        static void DrawDoubleWallRoom(dynamic modelSpace, double x, double y, double width, double height, double thickness)
        {
            DrawRectangle(modelSpace, x, y, width, height);
            DrawRectangle(modelSpace, x + thickness, y + thickness, width - (2 * thickness), height - (2 * thickness));
        }

        static void DrawRectangle(dynamic modelSpace, double x, double y, double width, double height)
        {
            double[] p1 = { x, y, 0 };
            double[] p2 = { x + width, y, 0 };
            double[] p3 = { x + width, y + height, 0 };
            double[] p4 = { x, y + height, 0 };

            modelSpace.AddLine(p1, p2);
            modelSpace.AddLine(p2, p3);
            modelSpace.AddLine(p3, p4);
            modelSpace.AddLine(p4, p1);
        }

        static void DrawDoorForRoom(dynamic modelSpace, double roomX, double roomY, double roomWidth, double roomHeight,
            double wallThickness, string wall, double offset, double doorWidth)
        {
            if (wall == "bottom")
            {
                double startX = roomX + offset;
                double endX = startX + doorWidth;
                double y = roomY;

                modelSpace.AddLine(
                    new double[] { roomX, y, 0 },
                    new double[] { startX, y, 0 });

                modelSpace.AddLine(
                    new double[] { endX, y, 0 },
                    new double[] { roomX + roomWidth, y, 0 });

                double[] hinge = { startX, y, 0 };
                modelSpace.AddLine(
                    hinge,
                    new double[] { startX, y + doorWidth, 0 });

                modelSpace.AddArc(hinge, doorWidth, 0, Math.PI / 2);
            }
            else if (wall == "top")
            {
                double startX = roomX + offset;
                double endX = startX + doorWidth;
                double y = roomY + roomHeight;

                modelSpace.AddLine(
                    new double[] { roomX, y, 0 },
                    new double[] { startX, y, 0 });

                modelSpace.AddLine(
                    new double[] { endX, y, 0 },
                    new double[] { roomX + roomWidth, y, 0 });

                double[] hinge = { startX, y, 0 };
                modelSpace.AddLine(
                    hinge,
                    new double[] { startX, y - doorWidth, 0 });

                modelSpace.AddArc(hinge, doorWidth, Math.PI * 1.5, 0);
            }
            else if (wall == "left")
            {
                double startY = roomY + offset;
                double endY = startY + doorWidth;
                double x = roomX;

                modelSpace.AddLine(
                    new double[] { x, roomY, 0 },
                    new double[] { x, startY, 0 });

                modelSpace.AddLine(
                    new double[] { x, endY, 0 },
                    new double[] { x, roomY + roomHeight, 0 });

                double[] hinge = { x, startY, 0 };
                modelSpace.AddLine(
                    hinge,
                    new double[] { x + doorWidth, startY, 0 });

                modelSpace.AddArc(hinge, doorWidth, Math.PI / 2, Math.PI);
            }
            else if (wall == "right")
            {
                double startY = roomY + offset;
                double endY = startY + doorWidth;
                double x = roomX + roomWidth;

                modelSpace.AddLine(
                    new double[] { x, roomY, 0 },
                    new double[] { x, startY, 0 });

                modelSpace.AddLine(
                    new double[] { x, endY, 0 },
                    new double[] { x, roomY + roomHeight, 0 });

                double[] hinge = { x, startY, 0 };
                modelSpace.AddLine(
                    hinge,
                    new double[] { x - doorWidth, startY, 0 });

                modelSpace.AddArc(hinge, doorWidth, 0, Math.PI / 2);
            }
        }

        static void DrawWindowForRoom(dynamic modelSpace, double roomX, double roomY, double roomWidth, double roomHeight,
            double wallThickness, string wall, double offset, double windowWidth)
        {
            if (wall == "top")
            {
                double startX = roomX + offset;
                double endX = startX + windowWidth;
                double yOuter = roomY + roomHeight;
                double yInner = yOuter - wallThickness;

                modelSpace.AddLine(
                    new double[] { startX, yOuter, 0 },
                    new double[] { endX, yOuter, 0 });

                modelSpace.AddLine(
                    new double[] { startX, yInner, 0 },
                    new double[] { endX, yInner, 0 });

                modelSpace.AddText("Window", new double[] { startX, yOuter + 200, 0 }, 140);
            }
        }

        static void DrawBed(dynamic modelSpace, double x, double y, double width, double height)
        {
            DrawRectangle(modelSpace, x, y, width, height);
            DrawRectangle(modelSpace, x + 80, y + height - 250, 500, 170);
            DrawRectangle(modelSpace, x + width - 580, y + height - 250, 500, 170);
            modelSpace.AddText("Bed", new double[] { x + width / 2 - 150, y + height / 2, 0 }, 140);
        }

        static void DrawTable(dynamic modelSpace, double x, double y, double width, double height)
        {
            DrawRectangle(modelSpace, x, y, width, height);
            modelSpace.AddCircle(new double[] { x + 80, y + 80, 0 }, 30);
            modelSpace.AddCircle(new double[] { x + width - 80, y + 80, 0 }, 30);
            modelSpace.AddCircle(new double[] { x + 80, y + height - 80, 0 }, 30);
            modelSpace.AddCircle(new double[] { x + width - 80, y + height - 80, 0 }, 30);
            modelSpace.AddText("Table", new double[] { x + width / 2 - 180, y + height / 2, 0 }, 120);
        }

        static void DrawToilet(dynamic modelSpace, double x, double y)
        {
            modelSpace.AddCircle(new double[] { x + 200, y + 500, 0 }, 180);
            DrawRectangle(modelSpace, x + 100, y, 200, 320);
            DrawRectangle(modelSpace, x, y + 700, 400, 220);
            modelSpace.AddText("WC", new double[] { x + 450, y + 450, 0 }, 120);
        }
    }
}