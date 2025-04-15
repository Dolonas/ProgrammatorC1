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
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class ProgrammatorC
	{
		private KompasObject _kompas;
		private IApplication _kompas7;
		private KompasDocument _activeDocument;
		private DocumentTypeEnum _activeDocumentType;
		private ksDocument2D _doc;
		private ksDocument3D _partOrAssembly;


		// Перестроить активный вид
		private void RebuildSelectedView()
		{
			if(_activeDocumentType == DocumentTypeEnum.ksDocumentDrawing)
				_doc.ksRebuildDocument();
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
		
		//Перестроить деталь или сборку
		private void Rebuild3dPart()
		{
			_activeDocument = (KompasDocument)_kompas7.ActiveDocument;
			_activeDocumentType = _activeDocument.DocumentType;
			// if (_activeDocumentType != DocumentTypeEnum.ksDocumentAssembly &&
			//     _activeDocumentType != DocumentTypeEnum.ksDocumentPart) return;
			_partOrAssembly = (IKompasDocument3D)_activeDocument;
			_kompas7.MessageBoxEx($"Дошёл до перестройки документа типа {_partOrAssembly.fileName}", "Информация ", 2);
			_partOrAssembly.RebuildDocument();
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
					result = "7-Перестроить деталь или сборку";
					command = 7;
					break;
				case 0:
					command = -1;
					itemType = 19; // "ENDMENU"
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
			_activeDocumentType = _activeDocument.DocumentType;
			
			if(_activeDocumentType == DocumentTypeEnum.ksDocumentDrawing)
			{
				_doc = (ksDocument2D)_kompas.ActiveDocument2D();
				if (_doc == null)
					return;
			}
			
			if(_activeDocumentType == DocumentTypeEnum.ksDocumentAssembly)
			{
				_partOrAssembly = (ksDocument3D)_kompas.Document3D();
				if (_partOrAssembly == null)
					return;
			}

			switch (command)
			{
				case 1:	RebuildSelectedView();			break; // перестроить вид
				case 2: ChangeSelectedLinesToTypeThink();			break; // Изменить тип линий на тонкую
				case 3: ChangeSelectedLinesToTypeInvisible();	break; // Изменить тип линий на вспомогательную
				case 4:	HideSelectedToInvisibleLayer();				break; // Убрать выделенное на скрытый слой
				case 5:	CleanUpRecordsOfChangesInAllSheets();				break; // Затирание извещений
				case 6:	CleanUpRecordsOfDates();				break; // Затирание дат
				case 7:	Rebuild3dPart();				break; //Перестроить деталь или сборку
			}

			//_kompas7.MessageBoxEx("Готово", "Информация ", 3);
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
	}

}
