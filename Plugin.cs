using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.IO;
using System.Net.Http;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using SystemVariableChangedEventArgs = Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs;
using SystemVariableChangedEventHandler = Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventHandler;

namespace AutoCADSimplePlugin
{
    public class ExampleRibbon : IExtensionApplication
    {
        public void Initialize()
        {
            BuildRibbonTab();
        }
        public void Terminate() { }
        void ComponentManager_ItemInitialized(object sender, Autodesk.Windows.RibbonItemEventArgs e)
        {
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                BuildRibbonTab();
                Autodesk.Windows.ComponentManager.ItemInitialized -=
                   ComponentManager_ItemInitialized;
            }
        }
        void BuildRibbonTab()
        {
            if (!IsLoaded())
            {
                CreateRibbonTab();
                Application.SystemVariableChanged += new SystemVariableChangedEventHandler(ACADApp_SystemVariableChanged);
            }
        }
        bool IsLoaded()
        {
            bool _loaded = false;
            RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
            foreach (RibbonTab tab in ribCntrl.Tabs)
            {
                if (tab.Id.Equals("ACAD.ID_TabHome"))
                {
                    if (tab.FindPanel("SimplePlugin") != null)
                    {
                        _loaded = true;
                        break;
                    }
                }
                else _loaded = false;
            }
            return _loaded;
        }
        void ACADApp_SystemVariableChanged(object sender, SystemVariableChangedEventArgs e)
        {
            if (e.Name.Equals("WSCURRENT")) BuildRibbonTab();
        }
        void CreateRibbonTab()
        {
            try
            {
                RibbonControl ribCntrl = Autodesk.Windows.ComponentManager.Ribbon;
                RibbonTab home = ribCntrl.FindTab("ACAD.ID_TabHome");
                AddContent(home);
                ribCntrl.UpdateLayout();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }
        void AddContent(RibbonTab ribTab)
        {
            try
            {
                RibbonPanelSource ribSourcePanel = new RibbonPanelSource();
                ribSourcePanel.Id = "SimplePlugin";
                ribSourcePanel.Title = "AutoCAD Simple Plugin";
                RibbonPanel ribPanel = new RibbonPanel();
                ribPanel.Source = ribSourcePanel;

                ribPanel.Source.Items.Clear();

                RibbonButton ribbonButton = new RibbonButton();
                ribbonButton.ShowText = true;
                ribbonButton.Text = "Execute";
                ribbonButton.LargeImage = LoadImage("fire_16");
                ribbonButton.ShowText = true;
                ribbonButton.ShowImage = true;
                ribbonButton.Size = RibbonItemSize.Large;
                ribbonButton.Orientation = System.Windows.Controls.Orientation.Vertical;
                ribbonButton.CommandHandler = new RibbonCommandHandler();

                ribSourcePanel.Items.Add(ribbonButton);
                ribTab.Panels.Add(ribPanel);
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.
                  DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.Message);
            }
        }

        System.Windows.Media.Imaging.BitmapImage LoadImage(string ImageName)
        {
            return new System.Windows.Media.Imaging.BitmapImage(
                new Uri("pack://application:,,,/AutoCADSimplePlugin;component/" + ImageName + ".png"));
        }

        class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }
            public event EventHandler CanExecuteChanged;
            public void Execute(object parameter)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (parameter is RibbonButton)
                {
                    RibbonButton button = parameter as RibbonButton;
                    Database database = doc.Database;
                    string ip;
                    try
                    {
                        using var client = new HttpClient();
                        ip = client.GetStringAsync("https://ipv4.icanhazip.com").Result.Trim();
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage("\nFailed to get IP address: " + ex.Message);
                        return;
                    }

                    var ptRes = doc.Editor.GetPoint("\nSpecify insertion point: ");
                    if (ptRes.Status != PromptStatus.OK) return;
                    Point3d pt = ptRes.Value;

                    using (doc.LockDocument())
                    {
                        using (var tr = database.TransactionManager.StartTransaction())
                        {
                            BlockTable acBlkTbl;
                            acBlkTbl = tr.GetObject(database.BlockTableId,
                                                         OpenMode.ForRead) as BlockTable;

                            BlockTableRecord acBlkTblRec;
                            acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                            OpenMode.ForWrite) as BlockTableRecord;

                            using (MText objText = new MText())
                            {
                                objText.Location = pt;

                                objText.Contents = ip;
                                objText.TextHeight = 16;
                                objText.Rotation = 90.5 * (Math.PI / 180.0);
                                objText.TextStyleId = database.Textstyle;

                                acBlkTblRec.AppendEntity(objText);

                                tr.AddNewlyCreatedDBObject(objText, true);
                            }

                            tr.Commit();
                        }
                    }

                    var progress = new ProgressWindow();
                    progress.Show();

                    progress.Dispatcher.InvokeAsync(async () =>
                    {
                        try
                        {
                            string dwgPath = Path.Combine("C:\\heavy.dwg");
                            if (!File.Exists(dwgPath))
                            {
                                doc.Editor.WriteMessage("\nFailed to open DWG.\nFile must be located in C:\\heavy.dwg");
                            }
                            Application.DocumentManager.Open(dwgPath, false);
                            progress.Focus();
                        }
                        catch (System.Exception ex)
                        {
                            doc.Editor.WriteMessage("\nFailed to open DWG: " + ex.Message);
                        }
                        finally
                        {
                            await System.Threading.Tasks.Task.Delay(3000);
                            progress.Dispatcher.Invoke(() => progress.Close());
                        }
                    });
                }
            }
        }
    }
}