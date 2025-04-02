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


		// Пересечь прямые
		private void RebuildSelectedView()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (arr != null)
			{
				doc.ksLine(10, 10, 0);
				doc.ksLine(15, 5, 90);
				mat.ksIntersectLinLin(10, 10, 0, 15, 5, 90, arr);
				DrawPointByArray(arr);
				arr.ksDeleteArray();
			}
		}


		// Пересечь кривые
		private void ChangeSelectedLinesToTypeThink()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			if (arr != null)
			{
				doc.ksBezier(0, 0);
					doc.ksPoint(20, 0, 0);
					doc.ksPoint(10, 20, 0);
					doc.ksPoint(20, 40, 0);
					doc.ksPoint(30, 20, 0);
					doc.ksPoint(20, 0, 0);
				int pp1 = doc.ksEndObj();

				doc.ksBezier(0, 0);
					doc.ksPoint(0, 20, 0);
					doc.ksPoint(20, 10, 0);
					doc.ksPoint(40, 20, 0);
					doc.ksPoint(20, 30, 0);
					doc.ksPoint(0, 20, 0);
				int pp2 = doc.ksEndObj();

				mat.ksIntersectCurvCurv(pp1, pp2, arr);
				DrawPointByArray(arr);
				arr.ksDeleteArray();
			}
		}


		// Пересечь отрезок и дугу
		private void ChangeSelectedLinesToTypeInviseble()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);

			if (arr != null)
			{
				doc.ksLineSeg(0, 40, 100, 40, 1);					// Отрисовка отрезка
				doc.ksArcByPoint(50, 40, 20, 30, 40, 70, 40, 1, 1);	// Отрисовка дуги по центру и конечным точкам
				double a1 = mat.ksAngle(50, 40, 30, 40);			// Начальный угол дуги
				double a2 = mat.ksAngle(50, 40, 70, 40);			// Конечный угол дуги
    
				// Получить координаты точек пересечения отрезка и дуги
				// Первая точка отрезка (0, 40), Вторая точка отрезка (100, 40),
				// Центр дуги (50, 40), Радиус дуги 20

				mat.ksIntersectLinSArc (0, 40, 100, 40, 50, 40, 20, a1, a2, 1, arr);
			    DrawPointByArray(arr);	// Отрисовка точек пересечения
    
				arr.ksDeleteArray();
			}
		}


		// Касательная из точки
		private void HideSelectedToInvisibleLayer()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

			if ((arr != null) & (par != null))
			{
				doc.ksPoint(10, 50, 3);			// Отрисовка точки
				doc.ksCircle(50, 10, 40, 1);	// Отрисовка окружности
    
				// Получить точки касания окружности и прямой, проходящей через заданную точку
				// Координаты внешней точки (10, 50), Координаты центра (50, 10),
				// радиус окружности 40

				mat.ksTanLinePointCircle(10, 50, 50, 10, 40, arr);
				DrawPointByArray(arr);		// Отрисовка точек пересечения
         
				// Отрисовка касательных
				for (int i = 0; i < arr.ksGetArrayCount(); i ++)
				{
					arr.ksGetArrayItem(i, par);	// Параметры текущей точки
					doc.ksLine(10, 50, mat.ksAngle(10, 50, par.x, par.y));
				}
      
				arr.ksDeleteArray();
			}
		}


		// Касательная под углом
		private void TanToAngle()
		{
			ksDynamicArray arr = (ksDynamicArray) kompas.GetDynamicArray(ldefin2d.POINT_ARR);
			ksMathPointParam par = (ksMathPointParam) kompas.GetParamStruct((short) StructType2DEnum.ko_MathPointParam);

			if ((arr != null) & (par != null))
			{
				doc.ksLineSeg(0, 40, 100, 40, 1);	// Отрисовка отрезка
				doc.ksCircle(50, 10, 40, 1);		// Отрисовка окружности
    
				// Получить точки касания окружности и прямой, проходящей под заданным углом
				// Координаты центра (50, 10), радиус окружности 40, Угол касательной прямой 45

				mat.ksTanLineAngCircle(50, 10, 40, 45, arr);
				DrawPointByArray(arr);			// Отрисовка точек пересечения
         
				// Отрисовка касательных
				for (int i = 0; i < arr.ksGetArrayCount(); i ++)
				{
					arr.ksGetArrayItem(i, par);	// Параметры текущей точки
					doc.ksLine(par.x, par.y, 45);
				}
       
			    arr.ksDeleteArray();
			}
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


		//Симметрия точки
		private void SymmetryPoint()
		{
			double x = 0;	// Результат симметрии точки
			double y = 0;

			doc.ksPoint(30, 60, 3);				// Отрисовка точки
			doc.ksLineSeg(0, 50, 60, 50, 3);	// Отрисовка отрезка
  
			// Получить координаты точки, симметричной относительно заданной оси
			mat.ksSymmetry(30, 60, 0, 50, 60, 50, out x, out y);
  
			doc.ksPoint(x, y, 5);	// Отрисовка результирующей точки
			// Результат симметрии
			kompas.ksMessage(string.Format("x = {0:.##} y = {1:.##}", x, y));
		}


		// Сопрягающие окружности к двум прямым
		private void CouplingTwoLines()
		{
			// Создать интерфейс массива координат точек сопряжения
			ksCON iCON = (ksCON) kompas.GetParamStruct((short) StructType2DEnum.ko_CON);

			if (iCON != null)	// Интерфейс создан
			{
				doc.ksLine(100, 100, 45);	// Отрисовка прямых - Первая прямая
				doc.ksLine(100, 100, -45);	// Вторая прямая
    
				// Получить параметры окружностей, касательной к двум прямым
				// Радиус сопряжения 20

				mat.ksCouplingLineLine(100, 100, 45, 100, 100, -45, 20, iCON);

				// Отрисовка сопрягающихся окружностей и точек касания
				for (int i = 0; i < iCON.GetCount(); i ++)
				{
					doc.ksCircle(iCON.GetXc(i), iCON.GetYc(i), 20, 2);
					doc.ksPoint(iCON.GetX1(i), iCON.GetY1(i), i);
					doc.ksPoint(iCON.GetX2(i), iCON.GetY2(i), i);
				}
  
				// Результат сопряжения
				string msg = string.Format("count = {0:.##} con[0].x1 = {1:.##} con[0].y1 = {2:.##} con[0].x2 = {3:.##} con[0].y2 = {4:.##} ...",
					iCON.GetCount(), 
					iCON.GetX1(0), 
					iCON.GetY1(0), 
					iCON.GetX2(0), 
					iCON.GetY2(0));
				kompas.ksMessage(msg);
			}
		}


		// Свойства вида
		private void RebuildDocument()
		{
			doc.ksRebuildDocument ();
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
