using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;

namespace RevitAPI_3_8_1
{
    public class MainViewViewModel
    {
        private ExternalCommandData _commandData;
        private Document _doc;
        public DelegateCommand ExportIFC { get;}
        public DelegateCommand ExportNWC { get; }
        public DelegateCommand ExportImmage { get; }
        public MainViewViewModel(ExternalCommandData commandData)        
        {
            _commandData = commandData;
            _doc = _commandData.Application.ActiveUIDocument.Document;
            ExportIFC = new DelegateCommand(OnExportIFC);
            ExportNWC = new DelegateCommand(OnExportNWC);
            ExportImmage = new DelegateCommand(OnExportImmage);
        }

        public Result OnExportIFC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            using (Transaction ts = new Transaction(doc, "Создание листов"))
            {
                ts.Start();
                var IfcOption = new IFCExportOptions();
                var saveDialog = new System.Windows.Forms.SaveFileDialog
                {
                    OverwritePrompt = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Filter = "All files (*.*)|*.*",
                    FileName = "Export.ifc",
                    DefaultExt = ".ifc"
                };
                string selectedFilePath = string.Empty;
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = saveDialog.FileName;
                }
                if (string.IsNullOrEmpty(selectedFilePath))
                    return Result.Cancelled;
                doc.Export(selectedFilePath, "Экспорт.ifc", IfcOption);                
                return Result.Succeeded;
                ts.Commit();
            }
            RaiseCloseRequest();
        }  

        public void OnExportNWC()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Создание листов"))
            {

                ts.Start();
                var nwcOption = new NavisworksExportOptions();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.nvc", nwcOption);
                ts.Commit();

            }
            RaiseCloseRequest();
        }

        public void OnExportImmage()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (var ts = new Transaction(doc, "Создание листов"))
            {

                ts.Start();
                ViewPlan viewPlan = new FilteredElementCollector(doc)
                                    .OfClass(typeof(ViewPlan))
                                    .Cast<ViewPlan>()
                                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan && v.Name.Equals("Level 1"));
                var imageOption = new ImageExportOptions
                {
                    ZoomType = ZoomFitType.FitToPage, 
                    PixelSize = 2024, 
                    FilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
                    FitDirection = FitDirectionType.Horizontal,
                    HLRandWFViewsFileType = ImageFileType.PNG, 
                    ShadowViewsFileType = ImageFileType.PNG, 
                    ImageResolution = ImageResolution.DPI_600, 
                    ExportRange = ExportRange.CurrentView                                                          // тот который мы отобрали с помощью FilteredElementCollectorю
                };
                doc.ExportImage(imageOption);
                ts.Commit();

            }
            RaiseCloseRequest();
        }

        public event EventHandler HideRequest;
        private void RaiseHideRequest()
        {
            HideRequest?.Invoke(this, EventArgs.Empty);
        }
        public event EventHandler ShowRequest;
        private void RaiseShowRequest()
        {
            ShowRequest?.Invoke(this, EventArgs.Empty);
        }
        public event EventHandler CloseRequest;
        private void RaiseCloseRequest()
        {
            CloseRequest?.Invoke(this, EventArgs.Empty);
        }
    }
}

