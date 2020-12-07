using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Crypto.Parameters;

namespace TestClient
{
    static class Extensions
    {
        public static string ToHashKey(this XSSFColor color)
        {
            return BitConverter.ToString(color.RGB ?? new byte[] {0,0,0});
        }

        public static (int row, int col) N(this (int row, int col) pos)
        {
            return (pos.row - 1, pos.col);
        }
        public static (int row, int col) S(this (int row, int col) pos)
        {
            return (pos.row + 1, pos.col);
        }
        public static (int row, int col) E(this (int row, int col) pos)
        {
            return (pos.row , pos.col +1);
        }
        public static (int row, int col) W(this (int row, int col) pos)
        {
            return (pos.row , pos.col -1);
        }
    }
    class Legend
    {
        public Dictionary<BorderStyle, string> walls = new Dictionary<BorderStyle, string>();
        public Dictionary<BorderStyle, string> doors = new Dictionary<BorderStyle, string>();
        public Dictionary<string, string> terrain = new Dictionary<string, string>();

        public Legend(ISheet legendSheet)
        {
            loadDoors(legendSheet);
            loadWalls(legendSheet);
            loadTerrain(legendSheet);
        }

        private void loadDoors(ISheet legendSheet)
        {
            var startCell = legendSheet.GetRow(1).GetCell(2);
            while (startCell != null && startCell.CellStyle.BorderRight != BorderStyle.None)
            {
                doors.Add(startCell.CellStyle.BorderRight, startCell.ToString());
                startCell = startCell.Sheet.GetRow(startCell.RowIndex + 1)?.GetCell(2);
            }
        }
        private void loadWalls(ISheet legendSheet)
        {
            var startCell = legendSheet.GetRow(1).GetCell(0);
            while (startCell != null && startCell.CellStyle.BorderRight != BorderStyle.None)
            {
                walls.Add(startCell.CellStyle.BorderRight, startCell.ToString());
                startCell = startCell.Sheet.GetRow(startCell.RowIndex + 1)?.GetCell(0);
            }
        }
        
        private void loadTerrain(ISheet legendSheet)
        {
            var toRetry = new List<(int Row, int Column)>();
            var startCell = legendSheet.GetRow(1).GetCell(4);
            while (startCell != null )
            {
                var foundColorCode = (startCell.CellStyle as XSSFCellStyle).FillForegroundXSSFColor.ToHashKey();
                var foundTerrain = startCell.ToString();
                if (terrain.ContainsKey(foundColorCode))
                {
                    // is this one wrong, or the one we put in before?
                    // maybe skip for now, and come back to read it again.
                    toRetry.Add((startCell.Address.Row, startCell.Address.Column));
                }
                else
                {
                   terrain.Add(foundColorCode, foundTerrain); 
                }
                
                startCell = startCell.Sheet.GetRow(startCell.RowIndex + 1)?.GetCell(4);
            }

            foreach (var address in toRetry)
            {
                startCell = legendSheet.GetRow(address.Row)?.GetCell(address.Column);
                var foundColorCode = (startCell.CellStyle as XSSFCellStyle).FillForegroundXSSFColor.ToHashKey();
                var foundTerrain = startCell.ToString();
                if (terrain.ContainsKey(foundColorCode))
                {
                    // is this one wrong, or the one we put in before?
                    // maybe skip for now, and come back to read it again.
                    //toRetry.Add(startCell.Address);
                }
                else
                {
                    terrain.Add(foundColorCode, foundTerrain); 
                }
            }
        }
    }
    
    class Program
    {

        static string PromptUser(string prompt)
        {
            Console.Write(prompt);
            return Console.ReadLine();
        }

        static int UserChoice(string question, string[] validAnswers)
        {
            while (true)
            {
                var answer = PromptUser(question);

                for (int i = 0; i < validAnswers.Length; i++)
                {
                    if (validAnswers[i].Equals(answer, StringComparison.CurrentCultureIgnoreCase))
                    {
                        return i;
                    }
                }

                Console.WriteLine($"Try typing: { String.Join( " or ", validAnswers) } ");
            }
        }

        static void displayWorldList()
        {
            var director = new DirectoryInfo("./");
            foreach (var file in director.EnumerateFiles("*.xlsx"))
            {
                if (file.Name.StartsWith("~$"))
                {
                    continue;
                }
                Console.WriteLine(file.Name.Replace(".xlsx",""));
            }
        }

        static XSSFWorkbook loadWorld(string filename)
        {
            using (var fs = new FileStream($"{filename}.xlsx", FileMode.Open, FileAccess.Read))
            {
                var wb = new XSSFWorkbook(fs);

                for (int i = 0; i < wb.Count; i++)
                {
                    Console.WriteLine($"Found Zone: {wb.GetSheetAt(i).SheetName}");  
                }

                return wb;
            }
        }
        
        static void Main(string[] args)
        {
            System.Console.Clear();
            System.Console.WriteLine("***My Bad ass text game***");

            displayWorldList();
            var worldName = PromptUser("Which world to play: ");
            var world = loadWorld(worldName);
            var legend = loadLegend(world);
            var position = startingPosition(world);
            if (position == null)
            {
                Console.WriteLine($"Error: no start found in {worldName}");    
            }

            
            var usersName = PromptUser("Type your name: ");
            
            
            
            var collected = new Dictionary<(int row, int col),CellInfo>();
            while (true)
            {
                var consolePos = Console.GetCursorPosition();
                Console.WriteLine("                                                                                                                    ");
                Console.WriteLine("                                                                                                                    ");
                PrintLocationDescription(position, legend, collected);
                Console.SetCursorPosition(consolePos.Left,consolePos.Top);
                var answerIndex = UserChoice("What do you want to do? ", new[] { "north", "south", "east", "west" });
                position = move(answerIndex, position, legend);    
            }
        }

        private static Legend loadLegend(XSSFWorkbook world)
        {
            var legendSheet = world.GetSheet("legend");
            
            return new Legend(legendSheet);
            
        }

        private static bool isPassable(BorderStyle border, Legend legend)
        {
            if (border == BorderStyle.None)
            {
                return true;
            }

            return legend.doors.ContainsKey(border);
        }

        private static ICell? move(in int answerIndex, ICell? currentPosition, Legend legend)
        {
           
            if( answerIndex == 0) // north
            {
                if (isPassable(currentPosition.CellStyle.BorderTop, legend))
                {
                    return currentPosition.Row.Sheet.GetRow(currentPosition.RowIndex-1).GetCell(currentPosition.ColumnIndex);
                }

                printWallInformation(legend, currentPosition.CellStyle.BorderTop, (currentPosition.CellStyle as XSSFCellStyle).TopBorderXSSFColor);
            }  
            if (answerIndex == 1) // south
            {
                if (isPassable(currentPosition.CellStyle.BorderBottom, legend))
                {
                    return currentPosition.Row.Sheet.GetRow(currentPosition.RowIndex + 1)
                        .GetCell(currentPosition.ColumnIndex);
                }

                printWallInformation(legend, currentPosition.CellStyle.BorderBottom, (currentPosition.CellStyle as XSSFCellStyle).BottomBorderXSSFColor);
            }
            if (answerIndex == 2)  // east
            {
                if (isPassable(currentPosition.CellStyle.BorderRight, legend))
                {
                    return currentPosition.Row.GetCell(currentPosition.ColumnIndex + 1);
                }
                printWallInformation(legend, currentPosition.CellStyle.BorderRight, (currentPosition.CellStyle as XSSFCellStyle).RightBorderXSSFColor);
            }
            if (answerIndex == 3)  // west
            {
                if (isPassable(currentPosition.CellStyle.BorderLeft, legend))
                {
                    return currentPosition.Row.GetCell(currentPosition.ColumnIndex - 1);
                }
                printWallInformation(legend, currentPosition.CellStyle.BorderLeft, (currentPosition.CellStyle as XSSFCellStyle).LeftBorderXSSFColor);
            }

            // NO MOVE
            return currentPosition;
        }

        private static void printWallInformation(Legend legend, BorderStyle obstruction, XSSFColor borderColor)
        {
            var wallType = legend.walls[obstruction];
            var wallMaterial = legend.terrain.ContainsKey(borderColor.ToHashKey()) ? legend.terrain[borderColor.ToHashKey()] : "default";

            Console.WriteLine($"You cannot walk though a {wallMaterial} {wallType} wall");
        }

        private static void PrintLocationDescription(ICell? position, Legend legend,
            Dictionary<(int row, int col),CellInfo> collected)
        {
            var comment = position.CellComment?.String;
            if (comment != null)
            {
                var justThePartAfterTheNotice =
                    comment.String.Split("Comment:", StringSplitOptions.RemoveEmptyEntries)[1];
                Console.WriteLine(justThePartAfterTheNotice);
            }

            printRoomDimensionInformation(position, legend, collected);
            printItemsInformation(position);
        }

        private static void printItemsInformation(ICell? position)
        {
            
        }


        
        public struct CellInfo
        {
            public (int row, int col) Position;
            public WallDoor WallDoorFlags;
            public Walls walls;
            public Doors doors;
            public readonly string TerrainKey;

            public CellInfo((int,int) position, string terrainKey)
            {
                Position = position;
                TerrainKey = terrainKey;
                walls = new Walls();
                doors = new Doors();
                WallDoorFlags = 0;
            }

            public bool AnyWallOrDoor()
            {
                return WallDoorFlags > 0;
            }

            public bool AnyWallDoorN()
            {
                return (WallDoorFlags & (WallDoor.DoorN | WallDoor.WallN)) > 0;
            }
            public bool AnyWallDoorS()
            {
                return (WallDoorFlags & (WallDoor.DoorS | WallDoor.WallS)) > 0;
            }
            public bool AnyWallDoorE()
            {
                return (WallDoorFlags & (WallDoor.DoorE | WallDoor.WallE)) > 0;
            }
            public bool AnyWallDoorW()
            {
                return (WallDoorFlags & (WallDoor.DoorW | WallDoor.WallW)) > 0;
            }
            
            [Flags]
            public enum WallDoor:byte
            {
                WallN = 0b00000001,
                WallS = 0b00000010,
                WallE = 0b00000100,
                WallW = 0b00001000,
                DoorN = 0b00010000,
                DoorS = 0b00100000,
                DoorE = 0b01000000,
                DoorW = 0b10000000,
            }
            public struct Walls
            {
                public Wall N;
                public Wall S;
                public Wall E;
                public Wall W;
            }

            public struct Doors
            {
                public Door N;
                public Door S;
                public Door E;
                public Door W;
            }
            public struct Wall
            {
                public BorderStyle wallKey;
                public string terrainKey;
            }

            public struct Door
            {
                public BorderStyle doorKey;
                public string terrainKey;
            }
        }
        
        private static CellInfo cellInfo(ICell? pos, Legend legend)
        {
            var cs = pos?.CellStyle as XSSFCellStyle;
            CellInfo ci = new CellInfo((pos.Address.Row, pos.Address.Column), cs?.FillForegroundXSSFColor?.ToHashKey());
            if (ci.TerrainKey == null)
            {
                return ci;
            }
            if (legend.walls.ContainsKey(cs.BorderTop))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.WallN;
                ci.walls.N.wallKey = cs.BorderTop;
                ci.walls.N.terrainKey = cs.TopBorderXSSFColor.ToHashKey();
            }
            if (legend.walls.ContainsKey(cs.BorderBottom))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.WallS;
                ci.walls.S.wallKey = cs.BorderBottom;
                ci.walls.S.terrainKey = cs.BottomBorderXSSFColor.ToHashKey();    
            }
            if (legend.walls.ContainsKey(cs.BorderLeft))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.WallW;
                ci.walls.W.wallKey = cs.BorderLeft;
                ci.walls.W.terrainKey = cs.LeftBorderXSSFColor.ToHashKey();    
            }
            if (legend.walls.ContainsKey(cs.BorderRight))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.WallE;
                ci.walls.E.wallKey = cs.BorderRight;
                ci.walls.E.terrainKey = cs.RightBorderXSSFColor.ToHashKey();    
            }
            // now do ci.doors
            if (legend.doors.ContainsKey(cs.BorderTop))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.DoorN;
                ci.doors.N.doorKey = cs.BorderTop;
                ci.doors.N.terrainKey = cs.TopBorderXSSFColor.ToHashKey();    
            }
            if (legend.doors.ContainsKey(cs.BorderBottom))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.DoorS;
                ci.doors.S.doorKey = cs.BorderBottom;
                ci.doors.S.terrainKey = cs.BottomBorderXSSFColor.ToHashKey();    
            }
            if (legend.doors.ContainsKey(cs.BorderLeft))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.DoorW;
                ci.doors.W.doorKey = cs.BorderLeft;
                ci.doors.W.terrainKey = cs.LeftBorderXSSFColor.ToHashKey();    
            }
            if (legend.doors.ContainsKey(cs.BorderRight))
            {
                ci.WallDoorFlags |= CellInfo.WallDoor.DoorE;
                ci.doors.E.doorKey = cs.BorderRight;
                ci.doors.E.terrainKey = cs.RightBorderXSSFColor.ToHashKey();    
            }

            return ci;
        }
        
        private static void printRoomDimensionInformation(ICell? position, Legend legend,Dictionary<(int row, int col),CellInfo> collected)
        {
            var currentPosition = position;

            var discoverQueue = new Queue<(int row, int col)>();
            var ci = cellInfo(currentPosition, legend); 

            bool AddLocation(CellInfo ci)
            {
                var pos = ci.Position;
                if (collected.ContainsKey(pos))
                {
                    return false;
                }
                collected.Add(pos, ci);

                var around = new[]
                {
                    ci.AnyWallDoorN() ? (pos.row,pos.col) : (pos.row - 1, pos.col), 
                    ci.AnyWallDoorS() ? (pos.row,pos.col) : (pos.row + 1, pos.col), 
                    ci.AnyWallDoorE() ? (pos.row,pos.col) : (pos.row, pos.col + 1),
                    ci.AnyWallDoorW() ? (pos.row,pos.col) : (pos.row, pos.col - 1)
                };
                foreach (var ar in around)
                {
                    if (collected.ContainsKey(ar) == false)
                    {
                        discoverQueue.Enqueue(ar);
                    }
                }

                return true;
            }

            if (AddLocation(ci))
            {
                while (discoverQueue.Count > 0)
                {
                    var nextRC = discoverQueue.Dequeue();
                    if (collected.ContainsKey(nextRC))
                    {
                        continue;
                    }
                    var cell = position.Sheet.GetRow(nextRC.row).GetCell(nextRC.col);
                    if (cell != null)
                    {
                        ci = cellInfo(cell, legend);
                        if (ci.TerrainKey != null)
                        {
                            if(!FixTouchingWalls(ci, collected))
                            {
                                AddLocation(ci);
                            }
                        }
                    }
                }
            }


            var roomWidth = collected.Values.Max(x => x.Position.col) - collected.Values.Min(x => x.Position.col) + 1;
            var roomHeight = collected.Values.Max(x => x.Position.row) - collected.Values.Min(x => x.Position.row) + 1;

            var size = roomSize(roomWidth, roomHeight);
            var shape = roomShape(roomWidth, roomHeight);

            Console.WriteLine($"{ position.Address } > You are in a {size} {shape}.");
           
            var sb1 = new StringBuilder(80);
            var sb2 = new StringBuilder(80);
            var sb3 = new StringBuilder(80);
            for (int r = position.RowIndex - 10; r < position.RowIndex + 10; r++)
            {
                 sb1.Clear();
                 sb2.Clear();
                 sb3.Clear();
                 for (int c = position.ColumnIndex -10; c < position.ColumnIndex + 10; c++) 
                 {
                     if (collected.ContainsKey((r, c)))
                     {
                         var cell = collected[(r, c)];

                         if (cell.AnyWallDoorW())
                         {
                             sb1.Append("┌");
                             sb2.Append("│");
                         }
                         else
                         {
                             sb1.Append(" ");
                             sb2.Append(" ");
                         }
                         if (cell.AnyWallDoorN())
                         {
                             sb1.Append("─");
                             sb2.Append(" ");
                         }else
                         {
                             sb1.Append(" ");
                             sb2.Append(" ");
                         }
                         if (cell.AnyWallDoorE())
                         {
                             sb1.Append("│");
                             sb2.Append(" ");
                         }
                         else
                         {
                             sb1.Append(" ");
                             sb2.Append(" ");
                         }
/*
    ┌─┬┐  ╔═╦╗  ╓─╥╖  ╒═╤╕
    │ ││  ║ ║║  ║ ║║  │ ││
    ├─┼┤  ╠═╬╣  ╟─╫╢  ╞═╪╡
    └─┴┘  ╚═╩╝  ╙─╨╜  ╘═╧╛
    ┌───────────────────┐
    │  ╔═══╗ Some Text  │▒
    │  ╚═╦═╝ in the box │▒
    ╞═╤══╩══╤═══════════╡▒
    │ ├──┬──┤           │▒
    │ └──┴──┘           │▒
    └───────────────────┘▒
     ▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒
 */



                     }
                     else
                     {
                         
                     }
                 }
            }
        }

        private static bool FixTouchingWalls(CellInfo ci, Dictionary<(int row, int col), CellInfo> collected)
        {
            if (ci.AnyWallOrDoor())
            {
                if (ci.AnyWallDoorN())
                {
                    var n = ci.Position.N();
                    if (collected.ContainsKey(n))
                    {
                        var c = collected[n];
                        if (c.Position == n && !c.AnyWallDoorS())
                        {
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.DoorN) > 0
                                ? CellInfo.WallDoor.DoorS
                                : 0;
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.WallN) > 0
                                ? CellInfo.WallDoor.WallS
                                : 0;
                            c.doors.S = ci.doors.N;
                            c.walls.S = ci.walls.N;
                            collected[n] = c;
                            return true;
                        }
                    }
                }

                if (ci.AnyWallDoorS())
                {
                    var s = ci.Position.S();
                    if (collected.ContainsKey(s))
                    {
                        var c = collected[s];
                        if (c.Position == s && !c.AnyWallDoorN())
                        {
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.DoorS) > 0
                                ? CellInfo.WallDoor.DoorN
                                : 0;
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.WallS) > 0
                                ? CellInfo.WallDoor.WallN
                                : 0;
                            c.doors.N = ci.doors.S;
                            c.walls.N = ci.walls.S;
                            collected[s] = c;
                            return true;
                        }
                    }
                }

                if (ci.AnyWallDoorW())
                {
                    var w = ci.Position.W();
                    if (collected.ContainsKey(w))
                    {
                        var c = collected[w];
                        if (c.Position == w && !c.AnyWallDoorE())
                        {
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.DoorW) > 0
                                ? CellInfo.WallDoor.DoorE
                                : 0;
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.WallW) > 0
                                ? CellInfo.WallDoor.WallE
                                : 0;
                            c.doors.E = ci.doors.W;
                            c.walls.E = ci.walls.W;
                            collected[w] = c;
                            return true;
                        }
                    }
                }

                if (ci.AnyWallDoorE())
                {
                    var e = ci.Position.E();
                    if (collected.ContainsKey(e))
                    {
                        var c = collected[e];
                        if (c.Position == e && !c.AnyWallDoorW())
                        {
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.DoorE) > 0
                                ? CellInfo.WallDoor.DoorW
                                : 0;
                            c.WallDoorFlags |= (ci.WallDoorFlags & CellInfo.WallDoor.WallE) > 0
                                ? CellInfo.WallDoor.WallW
                                : 0;
                            c.doors.W = ci.doors.E;
                            c.walls.W = ci.walls.E;
                            collected[e] = c;
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public static string roomShape(in int roomWidth, in int roomHeight)
        {
            if (Math.Min(roomHeight, roomWidth) == 1 && Math.Max(roomHeight, roomWidth) == 1)
            {
                return "room";
            }
            
            if (Math.Min(roomHeight, roomWidth) <= 2 && Math.Max(roomHeight, roomWidth) >= 8)
            {
                return "hallway";
            }
            
            if (Math.Max(roomHeight, roomWidth) >= 2)
            {
                return "room";
            }

            return $"area about {roomHeight * 5} by {roomWidth * 5}";

        }
        private static string roomSize(in int roomWidth, in int roomHeight)
        {
            var area = roomHeight * roomWidth;
            if (area * 5 < 50)
            {
                return "very small";
            }
            if (area * 5 < 100)
            {
                return "small";
            }
            if (area * 5 < 200)
            {
                return "medium";
            }
            if (area * 5 < 500)
            {
                return "large";
            }
            if (area * 5 < 1000)
            {
                return "huge";
            }
            
            return "giant";
        }

        private static ICell? startingPosition(XSSFWorkbook world)
        {
            var sheet = world.GetSheet("startzone");
            for (int rownum = 0; rownum < Int32.MaxValue; rownum++)
            {
                var row = sheet.GetRow(rownum);
                if (row != null)
                {
                    var foundCell = row.FirstOrDefault(c => c.StringCellValue == "start");
                    if (foundCell == null)
                    {
                        continue;
                    }

                    return foundCell;
                }
                else
                {
                    break;
                }
            }

            return null;
        }
        
    }
}