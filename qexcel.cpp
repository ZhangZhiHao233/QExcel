#include "qexcel.h"


QExcel::QExcel()
{
    pExcel = new QAxObject("Excel.Application");        //建立excel操作对象，并连接excel控件
    pExcel->dynamicCall("SetVisible(bool)", "true");   //设置不显示excel窗口
    pExcel->setProperty("DisplayAlerts", false);        //设置不显示提示信息
    pWorkBooks = pExcel->querySubObject("Workbooks");   //获取工作簿集合
	
	columnAll = 1;
	rowStart = 1;
	rowEnd = 1;
}


QAxObject* QExcel::CreateWorkBook(QAxObject* &sheet, QString& fileFullPath, const QString& firstSheetName)
{
   /* QString filePath;
    QDir dir;
    filePath = dir.currentPath();
    filePath.replace("/", "\\");
    filePath.append("\\");
    filePath += workBookName;

    QString fileFullPath = filePath + ".xlsx";*/

	QStringList list = fileFullPath.split(".");
	QString temp = list.at(0);
	QString workBookName = temp.section('/', -1);

	fileFullPath.replace("/", "\\");
	QFile file(fileFullPath);

    QAxObject* pWorkBook;
    QAxObject *pSheets;
    if(file.exists(fileFullPath))
    {
        if(!IsOpened(workBookName))//如果未打开
        {
            pWorkBooks->dynamicCall("Open(const QString&)", fileFullPath);  //打开工作簿
            pWorkBook = pExcel->querySubObject("ActiveWorkBook");        //获取当前活动工作簿
        }
        else
        {
            pWorkBook = GetOpenedWorkBook(workBookName);
        }

        pSheets = pWorkBook->querySubObject("WorkSheets");
        sheet = pSheets->querySubObject("Item(int)", 1);
    }
    else
    {
        pWorkBooks->dynamicCall("Add");
        pWorkBook = pExcel->querySubObject("ActiveWorkBook");
        pSheets = pWorkBook->querySubObject("WorkSheets");
        sheet = pSheets->querySubObject("Item(int)", 1);
        sheet->setProperty("Name", firstSheetName);
        pWorkBook->dynamicCall("SaveAs (const QString&)", QDir::toNativeSeparators(fileFullPath));
    }


    {
        openedWorkBook w;
        w.pWorkBook = pWorkBook;
        w.workBookName = workBookName;
        workBookList.append(w);
    }     

    return pWorkBook;
}


QAxObject* QExcel::OpenWorkBook(QString& fileFullPath)
{
    //QString filePath;
    //QDir dir;
    //filePath = dir.currentPath();
    //filePath.replace("/", "\\");
    //filePath.append("\\");
    //filePath += workBookName;

    //QString fileFullPath = filePath + ".xlsx";
	QStringList list = fileFullPath.split(".");
	QString temp = list.at(0);
	QString workBookName = temp.section('/', -1);

	fileFullPath.replace("/","\\");
    QFile file(fileFullPath);
    QAxObject* pWorkBook;

    if(file.exists(fileFullPath))
    {
        if(!IsOpened(workBookName))
        {
            pWorkBooks->dynamicCall("Open(const QString&)", fileFullPath);  //打开已有excel文件(工作簿集合)
            pWorkBook = pExcel->querySubObject("ActiveWorkBook");        //获取当前活动工作簿

            {
                openedWorkBook w;
                w.pWorkBook = pWorkBook;
                w.workBookName = workBookName;
                workBookList.append(w);
            }
        }
        else
        {
            pWorkBook = GetOpenedWorkBook(workBookName);
        }

    }
    else
    {
        pWorkBook = NULL;
    }

    return pWorkBook;
}


bool QExcel::CloseWorkBook(QAxObject* &workBook, const QString& workBookName)
{
    if(!workBook)
    {
       return false;
    }

    workBook->dynamicCall("Close(Boolean)", false);
	workBook = 0;
    for(int i = 0; i < workBookList.size(); i++)
    {
        if(!workBookName.compare(workBookList.at(i).workBookName))
        {
            delete workBookList.at(i).pWorkBook;
            workBookList.removeAt(i);
        }
        return true;
    }

    return false;
}


bool QExcel::DelWorkBook(const QString& workBookName)
{

        QString filePath;
        QDir dir;
        filePath = dir.currentPath();
        filePath.replace("/", "\\");
        filePath.append("\\");
        filePath += workBookName;

        QString fileFullPath = filePath + ".xlsx";
		if(QFile::exists(fileFullPath))
		{
			 QFile::remove(fileFullPath);
			 return true;
		 }
		else
		    return false;
}


QAxObject* QExcel::AddSheet(QAxObject* workbook, const QString& sheetName)
{
    QAxObject *pSheets = workbook->querySubObject("WorkSheets");

    QAxObject *temp =  IsSheetExist(pSheets, sheetName);
    if(temp)
    {
        return temp;
    }

    pSheets->querySubObject("Add()");
    QAxObject *add = pSheets->querySubObject("Item(int)", 1);
    add->setProperty("Name", sheetName);
    workbook->dynamicCall("Save()");
    QAxObject* pSheet = pSheets->querySubObject("Item(int)", 1);
    return pSheet;
}


QAxObject* QExcel::SelectSheet(QAxObject* workBook, const QString& sheetName)
{
    QAxObject *pSheets = workBook->querySubObject("WorkSheets");
    QAxObject *pSheet = pSheets->querySubObject("Item(const QString&)", sheetName);
    return pSheet;
}


bool QExcel::DelSheet(QAxObject* workbook, const QString& sheetName)
{
    QAxObject *pSheets = workbook->querySubObject("WorkSheets");
    QAxObject *sheet = IsSheetExist(pSheets, sheetName);
    if(!sheet)
    {
        return false;
    }
    else
    {
        sheet->dynamicCall("delete");
        workbook->dynamicCall("Save()");
        return true;
    }
}


void QExcel::InsertTitle(QAxObject* workBook, QAxObject* sheet, const QString& titleName,int mode)
{
	int topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn;
	GetUsageRange(sheet, &topLeftRow, &topLeftColumn, &bottomRightRow, &bottomRightColumn);

	int column = (mode == 1) ? 7 : 5;
	columnAll = column;
	QAxObject *range = 0;
	if((bottomRightRow == 1) && (bottomRightColumn == 1))//如果是首行的话直接插入
	{
		range = MergeCells(sheet, 1, 1, 1, column);
	}
	else//否则空两行再插入
	{
		range = MergeCells(sheet, bottomRightRow + 3, 1, bottomRightRow + 3, column);
		rowStart = bottomRightRow + 3;
	}

	SetCell(*range, QColor(255, 127, 39), QColor(0, 0, 0), 24, 300, 16, true);
	range->setProperty("Value", titleName);

	workBook->dynamicCall("Save()");
}


void QExcel::InsertTableTitle(QAxObject* workBook, QAxObject* sheet, int mode)
{
	int topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn;
	GetUsageRange(sheet, &topLeftRow, &topLeftColumn, &bottomRightRow, &bottomRightColumn);

	if(mode)
	{
		QAxObject *cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 1);
		QAxObject *font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("序号"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 2);

		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 10);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("姓名"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 3);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 15);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数一"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 4);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 12);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数二"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 5);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 12);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数三"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 6);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数四"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 7);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("排名"));
	}
	else
	{
		QAxObject *cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 1);
		QAxObject *font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("序号"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 2);

		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 10);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数一"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 3);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数二"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 4);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("参数三"));

		cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 5);
		font = cell->querySubObject("Font");
		font->setProperty("Bold", true);
		cell->setProperty("ColumnWidth", 8);
		cell->setProperty("HorizontalAlignment", -4108);
		cell->setProperty("Value", QStringLiteral("排名"));
	}

	workBook->dynamicCall("Save()");
}

void QExcel::InsertInfo(QAxObject* workBook, QAxObject* sheet, int count, int targetPoint, int time, int dist, 
	int rank)
{
	int topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn;
	GetUsageRange(sheet, &topLeftRow, &topLeftColumn, &bottomRightRow, &bottomRightColumn);
	QAxObject *cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 1);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(count));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 2);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(targetPoint));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 3);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(time));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 4);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(dist));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 5);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(rank));

	rowEnd = bottomRightRow + 1;
	workBook->dynamicCall("Save()");
}

void QExcel::InsertInfo(QAxObject* workBook, QAxObject* sheet, int count, const QString& userName, int serTargetCnt, 
	int serTime, int serDis, const QString&  grade, int rank)
{
	int topLeftRow, topLeftColumn, bottomRightRow, bottomRightColumn;
	GetUsageRange(sheet, &topLeftRow, &topLeftColumn, &bottomRightRow, &bottomRightColumn);
	QAxObject *cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 1);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(count));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 2);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", userName);

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 3);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(serTargetCnt));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 4);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(serTime));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 5);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(serDis));

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 6);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", grade);

	cell =  sheet->querySubObject("Cells(int, int)", bottomRightRow + 1, 7);
	cell->setProperty("HorizontalAlignment", -4108);
	cell->setProperty("Value", QString::number(rank));

	rowEnd = bottomRightRow + 1;
	workBook->dynamicCall("Save()");
}



void QExcel::CloseExcel()
{
    CloseAllWorkBook();
    pExcel->dynamicCall("Quit(void)");

    delete pWorkBooks;
    delete pExcel;

    pWorkBooks = 0;
    pExcel = 0;
}


bool QExcel::IsOpened(const QString& workBookName)
{
    for(int i = 0; i < workBookList.size(); i++)
    {
        if(!workBookName.compare(workBookList.at(i).workBookName))
        {
            return true;
        }
    }
    return false;
}


QAxObject* QExcel::IsSheetExist(QAxObject* pSheets, const QString& sheetName)
{
    int iCount = pSheets->property("Count").toInt();

    QAxObject *sheet;
    for(int i = 1; i <= iCount; i++)
    {
       sheet  = pSheets->querySubObject("Item(int)", i);
       QString name = sheet->property("Name").toString();
       if(!name.compare(sheetName,Qt::CaseSensitive))
       {
           return sheet;
       }
    }

    return NULL;
}


QAxObject* QExcel::GetOpenedWorkBook(const QString& workBookName)
{
    for(int i = 0; i < workBookList.size(); i++)
    {
        if(!workBookName.compare(workBookList.at(i).workBookName, Qt::CaseSensitive))
        {
            return workBookList.at(i).pWorkBook;
        }
    }

    return NULL;
}

QAxObject* QExcel::MergeCells(QAxObject* sheet, int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    //将需要合并的单元格范围整合成类似“A1:B10”的格式
    QString cell;
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(QString(":"));
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("HorizontalAlignment", -4108); //设置水平方向居中 -4131左对齐 -4152右对齐
    range->setProperty("VerticalAlignment", -4108);   //设置垂直方向居中 -4160上对齐 -4107下对齐
    range->setProperty("MergeCells", true);

    return range;
}


void QExcel::GetUsageRange(QAxObject* sheet, int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject *rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject *columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}


void QExcel::SetCell(QAxObject& cell, QColor& bgColor, QColor& txtColor, int height, int weight,
                       int fontSize, bool isBold)
{
    QAxObject *font = cell.querySubObject("Font");
    font->setProperty("Bold", isBold);
    font->setProperty("Size", fontSize);
    font->setProperty("Color", txtColor);

    cell.setProperty("RowHeight", height);
    cell.setProperty("WrapText", true);

    //只有单元格才能设置列宽
    {
        QAxObject *rows = cell.querySubObject("Rows");
        int row = rows->property("Count").toInt();

        QAxObject *columns = cell.querySubObject("Columns");
        int column = columns->property("Count").toInt();

        //qDebug()<<row<<column;

        if((row == 1) && (column == 1))
        {
            cell.setProperty("ColumnWidth", weight);
        }

    }

    QAxObject *interior = cell.querySubObject("Interior");
    interior->setProperty("Color", bgColor);
}


void QExcel::CloseAllWorkBook()
{
    for(int i = 0; i < workBookList.size();)
    {
        workBookList.at(i).pWorkBook->dynamicCall("Close(Boolean)", false);
        delete workBookList.at(i).pWorkBook;
        workBookList.removeAt(i);
        i = 0;
    }
}


void QExcel::SetBorder(QAxObject* workBook, QAxObject* sheet)
{
	QString cell;
	cell.append(QChar(1 - 1 + 'A'));
	cell.append(QString::number(rowStart));
	cell.append(QString(":"));
	cell.append(QChar(columnAll - 1 + 'A'));
	cell.append(QString::number(rowEnd));

	QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
	QAxObject* border = range->querySubObject("Borders");
	border->setProperty("Color", QColor(0, 0, 0));
	workBook->dynamicCall("Save()");
}