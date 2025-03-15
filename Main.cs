using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Module_6_5
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {

                List<Room> rooms = GetRooms(doc);

                if (rooms.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "В проекте нет помещений!");
                    return Result.Cancelled;
                }


                List<Room> sortedRooms = SortRooms(rooms);


                using (Transaction t = new Transaction(doc, "Автонумерация помещений"))
                {
                    t.Start();

                    int counter = 1;
                    foreach (Room room in sortedRooms)
                    {
                        SetRoomNumber(room, counter.ToString());
                        counter++;
                    }

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private List<Room> GetRooms(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(room => room.Location is LocationPoint)
                .ToList();
        }

        private List<Room> SortRooms(List<Room> rooms)
        {
            return rooms.OrderBy(room => room.Level.Elevation)
                .ThenBy(room => ((LocationPoint)room.Location).Point.X)
                .ThenBy(room => ((LocationPoint)room.Location).Point.Y)
                .ToList();
        }

        private void SetRoomNumber(Room room, string number)
        {
            Parameter numberParam = room.get_Parameter(BuiltInParameter.ROOM_NUMBER);
            if (numberParam != null && !numberParam.IsReadOnly)
            {
                if (numberParam.StorageType == StorageType.String)
                {
                    numberParam.Set(number);
                }
            }
        }
    }
}
