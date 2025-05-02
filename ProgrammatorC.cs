using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KAPITypes;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;

namespace ProgrammatorC
{
	enum DocType { Drawing, Part, Assembly, Spec, Fragment, Textual, Unnkown }
	
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class ProgrammatorC
	{
		private KompasObject _kompas;
		private IApplication _kompas7;
		private KompasDocument _activeDocument;
		private DocType _activeDocumentType;
		private ksDocument2D _doc;
		//private ksDocument3D _partOrAssembly;


		// Перестроить активный вид
		private void RebuildActiveElements()
		{
			AssignActiveDocumentType();
			if(_activeDocumentType == DocType.Drawing)
				_doc.ksRebuildDocument();
			if(_activeDocumentType == DocType.Assembly || _activeDocumentType == DocType.Part)
			{
				IKompasDocument3D partOrAssembly = (IKompasDocument3D)_kompas7.ActiveDocument;
				partOrAssembly.RebuildDocument();
			}
		}


		// Изменить тип линий на тонкую
		private void ChangeSelectedLinesToTypeThink()
		{
			
		}


		// Изменить тип линий на вспомогательную
		private void ChangeSelectedLinesToTypeInvisible()
		{
			
		}


		// Убрать выделенное на скрытый слой
		private void HideSelectedToInvisibleLayer()
		{
			AssignActiveDocumentType();
			if(_activeDocumentType != DocType.Drawing)
				return;
			_doc = (ksDocument2D)_kompas.ActiveDocument2D();
			if (_doc == null)
				return;

			var obj = _doc.ksLayer(1);
			ksLayerParam layerParam =
				(ksLayerParam)_kompas.GetParamStruct((short)StructType2DEnum.ko_LayerParam);
			layerParam.Init();
			layerParam.color = 0;
			layerParam.name = "Скрытые";
			layerParam.state = 2;
			_doc.ksSetObjParam(obj, layerParam, ldefin2d.VIEW_LAYER_STATE);
		}

		// Затирание извещений
		private void CleanUpRecordsOfChangesInAllSheets()
		{
			var sheetsNum = _doc.ksGetDocumentPagesCount();
			var cellList = new List<int>{140, 150, 160, 170, 180}; 
			for (var i = 0; i <= sheetsNum; i++)
			{
				var stamp = (ksStamp)_doc.GetStampEx(i);
				foreach (var cell in cellList)
				{
					stamp.ksClearStamp(cell);
				}
			}
		}
		
		//Затирание дат
		private void CleanUpRecordsOfDates()
		{
			var sheetsNum = _doc.ksGetDocumentPagesCount();
			var cellList = new List<int>{181, 180, 130, 131, 132, 133, 134, 135}; 
			for (var i = 0; i <= sheetsNum; i++)
			{
				var stamp = (ksStamp)_doc.GetStampEx(i);
				foreach (var cell in cellList)
				{
					try
					{
						stamp.ksClearStamp(cell);
					}
					catch (Exception e)
					{
						Console.WriteLine(e);
						continue;
					}
					
				}
			}
		}
		

		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Programmator - Создание Панели инструментов";
		}

		[return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "1-Перестроить вид";
					command = 1;
					break;
				case 2:
					result = "2-Изменить тип линий на тонкую";
					command = 2;
					break;
				case 3:
					result = "3-Изменить тип линий на вспомогательную";
					command = 3;
					break;
				case 4:
					result = "4-Убрать выделенное на скрытый слой";
					command = 4;
					break;
				case 5:
					result = "5-Затирание извещений";
					command = 5;
					break;
				case 6:
					result = "6-Затирание дат";
					command = 6;
					break;
				case 7:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
            return result;
		}

	
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas)
		{
			_kompas = (KompasObject) kompas;
			if (_kompas == null)
				return;
			
			_kompas7 = _kompas.ksGetApplication7() as IApplication;
			if (_kompas7 == null)
				return;

			_activeDocument = (KompasDocument)_kompas7.ActiveDocument;
			AssignActiveDocumentType();
			
			if(_activeDocumentType == DocType.Drawing)
			{
				_doc = (ksDocument2D)_kompas.ActiveDocument2D();
				if (_doc == null)
					return;
			}

			switch (command)
			{
				case 1:	RebuildActiveElements();			break; // перестроить вид
				case 2: ChangeSelectedLinesToTypeThink();			break; // Изменить тип линий на тонкую
				case 3: ChangeSelectedLinesToTypeInvisible();	break; // Изменить тип линий на вспомогательную
				case 4:	HideSelectedToInvisibleLayer();				break; // Убрать выделенное на скрытый слой
				case 5:	CleanUpRecordsOfChangesInAllSheets();				break; // Затирание извещений
				case 6:	CleanUpRecordsOfDates();				break; // Затирание дат
				//case 7:	Rebuild3dPart();				break; //Перестроить деталь или сборку
			}
		}


		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
		}


		public int ExternalGetToolBarId(short barType, short index)
		{
			var result = 0;

			if (barType == 0)
			{
				result = -1;
			}
			else
			{
				switch (index)
				{
					case 1:
						result = 3001;
						break;
					case 2:
						result = -1;
						break;
				}
			}

			return result;
		}

		#region COM Registration
		
		[ComRegisterFunction]
		public static void RegisterKompasLib(Type t)
		{
			try
			{
				var regKey = Registry.LocalMachine;
				var keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
				regKey = regKey.OpenSubKey(keyName, true);
				if (regKey == null) return;
				regKey.CreateSubKey("Kompas_Library");
				regKey = regKey.OpenSubKey("InprocServer32", true);
				if (regKey == null) return;
				regKey.SetValue(null,
					System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
				regKey.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"При регистрации класса для COM-Interop произошла ошибка:\n{ex}");
			}
		}
		
		
		[ComUnregisterFunction]
		public static void UnregisterKompasLib(Type t)
		{
			var regKey = Registry.LocalMachine;
			var keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
			var subKey = regKey.OpenSubKey(keyName, true);
			if (subKey == null) return;
			subKey.DeleteSubKey("Kompas_Library");
			subKey.Close();
		}
		#endregion

		private void AssignActiveDocumentType()
		{
			KompasDocument activeDocument = (KompasDocument)_kompas7.ActiveDocument;
			switch (activeDocument.DocumentType)
			{
				case DocumentTypeEnum.ksDocumentDrawing:
					_activeDocumentType = DocType.Drawing;
					break;
				case DocumentTypeEnum.ksDocumentAssembly:
					_activeDocumentType = DocType.Assembly;
					break;
				case DocumentTypeEnum.ksDocumentPart:
					_activeDocumentType = DocType.Part;
					break;
				case DocumentTypeEnum.ksDocumentSpecification:
					_activeDocumentType = DocType.Spec;
					break;
				case DocumentTypeEnum.ksDocumentFragment:
					_activeDocumentType = DocType.Fragment;
					break;
				case DocumentTypeEnum.ksDocumentTextual:
					_activeDocumentType = DocType.Textual;
					break;
				case DocumentTypeEnum.ksDocumentUnknown:
					_activeDocumentType = DocType.Unnkown;
					break;
				case DocumentTypeEnum.ksDocumentTechnologyAssembly:
					_activeDocumentType = DocType.Unnkown;
					break;
				default:
					_activeDocumentType = DocType.Unnkown;
					break;
			}
		}
	}
}
