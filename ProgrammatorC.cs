using Kompas6API5;
using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using KAPITypes;
using Kompas6Constants;

namespace ProgrammatorC
{
	// Класс ProgrammatorC - Программатор
	
	// 1. Пересечь прямые                      - Intersect2Line
	// 2. Пересечь кривые                      - Intersect2Curve
	// 3. Пересечь отрезок и дугу              - IntersectLineSegArc
	// 4. Касательная из точки                 - TanLinePointCircle
	// 5. Касательная под углом                - TanLineAngCircle
	// 6. Поворот точки                        - RotatePoint
	// 7. Симметрия точки                      - SymmetryPoint
	// 8. Сопрягающие окружности к двум прямым - Couplin2Lines
	// 9. Перепендикуляр                       - Perpendicular

	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class Step2New2
	{
		private KompasObject kompas;
		private ksDocument2D doc;
		private ksMathematic2D mat;

        /// <summary>
        /// Отрисовка точек пересечения в документе doc по
        /// присланному массиву и выдача пользователю их координат
        /// </summary>
        private void DrawPointByArray(ksDynamicArray arr)
		{
			if (arr != null)
			{
				// Создать интерфейс параметров математической точки
				ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

				if (par != null)
				{
					// Интерфейс создан
					for (int i = 0; i < arr.ksGetArrayCount(); i ++)
					{
						arr.ksGetArrayItem(i, par);
						doc.ksPoint(par.x, par.y, 5);
						string buf = string.Format("x = {0:.##} y = {1:.##}", par.x, par.y);
						kompas.ksMessage(buf);
					}
				}
			}
		}


		// Перестроить активный вид
		private void RebuildSelectedView()
		{
			doc.ksRebuildDocument();
		}


		// Пересечь кривые
		private void ChangeSelectedLinesToTypeThink()
		{
			
		}


		// Пересечь отрезок и дугу
		private void ChangeSelectedLinesToTypeInviseble()
		{
			
		}


		// Касательная из точки
		private void HideSelectedToInvisibleLayer()
		{
			
		}


		// Касательная под углом
		private void TanToAngle()
		{
			
		}


		// стирание извещений
		private void CleanUpRecordsOfChangesInAllSheets()
		{
			var sheetsNum = doc.ksGetDocumentPagesCount();
			var cellList = new List<int>{140, 150, 160, 170, 180}; 
			for (var i = 0; i <= sheetsNum; i++)
			{
				var stamp = (ksStamp)doc.GetStampEx(i);
				foreach (var cell in cellList)
				{
					stamp.ksClearStamp(cell);
				}
			}
		}
		private void CleanUpRecordsOfDates()
		{
			
		}

		[return: MarshalAs(UnmanagedType.BStr)] public string GetLibraryName()
		{
			return "Step2New2 - Создание Панели инструментов";
		}


		[return: MarshalAs(UnmanagedType.BStr)] public string ExternalMenuItem(short number, ref short itemType, ref short command)
		{
			string result = string.Empty;
			itemType = 1; // "MENUITEM"
			switch (number)
			{
				case 1:
					result = "1-Пересечь прямые";
					command = 1;
					break;
				case 2:
					result = "2-Пересечь кривые";
					command = 2;
					break;
				case 3:
					result = "3-Пересечь отрезок и дугу";
					command = 3;
					break;
				case 4:
					result = "4-Касательная из точки";
					command = 4;
					break;
				case 5:
					result = "5-Касательная под углом";
					command = 5;
					break;
				case 6:
					result = "6-затирание извещений";
					command = 6;
					break;
				case 7:
					result = "7-Симметрия точки";
					command = 7;
					break;
				case 8:
					result = "8-Сопрягающие окружности к двум прямым";
					command = 8;
					break;
				case 9:
					result = "9-Перестроить документ";
					command = 9;
					break;
				case 10:
					command = -1;
					itemType = 3; // "ENDMENU"
					break;
			}
            return result;
		}

	
		public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
		{
			kompas = (KompasObject) kompas_;
			if (kompas == null)
				return;

			doc = (ksDocument2D) kompas.ActiveDocument2D();
			if (doc == null)
				return;

			mat = (ksMathematic2D) kompas.GetMathematic2D();
			if (mat == null)
				return;

			switch (command)
			{
				case 1:	RebuildSelectedView();			break; // перестроить вид
				case 2: ChangeSelectedLinesToTypeThink();			break; // Изменить тип линий на тонкую
				case 3: ChangeSelectedLinesToTypeInviseble();	break; // Изменить тип линий на вспомогательную
				case 4:	HideSelectedToInvisibleLayer();				break; // Убрать выделенное на скрытый слой
				case 5:	CleanUpRecordsOfChangesInAllSheets();				break; // Затирание извещений
				case 6:	CleanUpRecordsOfDates();				break; // Затирание дат
			}

			kompas.ksMessageBoxResult();
		}


		public object ExternalGetResourceModule()
		{
			return Assembly.GetExecutingAssembly().Location;
		}


		public int ExternalGetToolBarId(short barType, short index)
		{
			int result = 0;

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
				RegistryKey regKey = Registry.LocalMachine;
				string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
				regKey = regKey.OpenSubKey(keyName, true);
				regKey.CreateSubKey("Kompas_Library");
				regKey = regKey.OpenSubKey("InprocServer32", true);
				regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
				regKey.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
			}
		}
		
		
		[ComUnregisterFunction]
		public static void UnregisterKompasLib(Type t)
		{
			RegistryKey regKey = Registry.LocalMachine;
			string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
			RegistryKey subKey = regKey.OpenSubKey(keyName, true);
			subKey.DeleteSubKey("Kompas_Library");
			subKey.Close();
		}
		#endregion
	}

}
