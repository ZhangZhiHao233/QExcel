/**
 * @file  qexcel.h
 * @brief 包含操作Excel的QExcel类的定义
 *
 * 该类通过COM来操作excel，主要使用QAxObject以及Excel VBA，须在.pro文件中添加 “QT += axcontainer”
 *
 * Excel的层次结构为 excel应用程序->工作簿->工作表->单元格
 * 一个类对象可以创建多个工作簿，每个工作簿可以操作多张工作表
 *
 * 在新线程中使用要单独初始化COM,OleInitialize(0);
 *
 * 使用流程：创建类对象->创建/打开工作簿->添加/打开/删除工作表->添加表信息->关闭工作簿->删除工作簿->关闭excel
 *
 * @author zzh
 * @version 1.0
 * @date 2018.8
*/

#ifndef QEXCEL_H
#define QEXCEL_H

#include <QAxObject>
#include <QString>
#include <QDir>
#include <QFile>
#include <QColor>
#include <QDebug>
#include <QList>


struct openedWorkBook
{
    QAxObject *pWorkBook;//工作簿对象指针
    QString workBookName;//工作簿名称
};


class QExcel
{
public:
    QExcel();//打开excel进程,获取所有工作簿

private:
    QAxObject *pExcel;     //excel操作对象指针
    QAxObject *pWorkBooks; //工作簿集合指针
    QList<openedWorkBook> workBookList;//保存已经打开的工作簿名称以及对象指针

	int columnAll;//使用列数，用于设置边框
	int rowStart;//起始行
	int rowEnd;//结束行
public:
    /*********************************************************
            FunctionName:   CreateWorkBook
            Purpose:        创建工作簿
            Parameter:
                            1 sheet [QAxObject* &, OUT]
                               第一张工作表指针
                            2 fileFullPath [QString&, IN]
                               工作簿完整路径
                            3 firstSheetName [const QString&, IN]
                               第一张表名称

            Return:
                            如果该工作簿存在，则打开并返回其指针
                            如果该工作簿不存在，则返回创建的工作簿指针
            Remark:
                            创建工作簿时必须默认创建一张表
        ***********************************************************/
    QAxObject* CreateWorkBook(QAxObject* &sheet, QString& fileFullPath, const QString& firstSheetName);


    /*********************************************************
            FunctionName:   OpenWorkBook
            Purpose:        打开工作簿
            Parameter:
                            1 fileFullPath [QString&, IN]
                               工作簿完整路径
            Return:
                            如果该工作簿存在，则打开并返回其指针
                            如果该工作簿不存在，则返回NULL
            Remark:
        ***********************************************************/
    QAxObject* OpenWorkBook(QString& fileFullPath);


    /*********************************************************
            FunctionName:   CloseWorkBook
            Purpose:        关闭工作簿
            Parameter:
                            1 workBook [QAxObject* &, IN]
                               工作簿指针
                            2 workBookName [const QString&, IN]
                               工作簿名称
            Return:         成功返回true，失败返回false
            Remark:
        ***********************************************************/
    bool CloseWorkBook(QAxObject* &workBook, const QString& workBookName);


    /*********************************************************
           FunctionName:   DelWorkBook
           Purpose:        删除工作簿
           Parameter:
                           1 workBookName [const QString&, IN]
                              工作簿名称
           Return:         成功返回true，失败返回false
           Remark:
       ***********************************************************/
    bool DelWorkBook(const QString& workBookName);


    /*********************************************************
           FunctionName:   AddSheet
           Purpose:        添加工作表
           Parameter:
                           1 workBook [QAxObject* &, IN]
                              工作簿指针

                           2 sheetName [const QString&, IN]
                              工作表名称
           Return:
                           如果该工作表存在，则返回其指针
                           如果该工作表不存在，则返回创建的工作表指针
           Remark:
       ***********************************************************/
    QAxObject* AddSheet(QAxObject* workbook, const QString& sheetName);


    /*********************************************************

           FunctionName:   SelectSheet
           Purpose:        选择工作表
           Parameter:
                           1 workBook [QAxObject* &, IN]
                              工作簿指针

                           2 sheetName [const QString&, IN]
                              工作表名称
           Return:
                           如果该工作表存在，则返回其指针
                           如果该工作表不存在，则返回NULL
           Remark:
       ***********************************************************/
    QAxObject* SelectSheet(QAxObject* workBook, const QString& sheetName);


    /*********************************************************
            FunctionName:   DelSheet
            Purpose:        删除工作表
            Parameter:
                            1 workBook [QAxObject* &, IN]
                               工作簿指针

                            2 sheetName [const QString&, IN]
                               工作表名称
            Return:         成功返回true，失败返回false
            Remark:
                            一个工作簿至少包含一个工作表，因此不能删除最后一张表
                            可以直接删除工作簿
        ***********************************************************/
    bool DelSheet(QAxObject* workbook, const QString& sheetName);


    /*********************************************************
            FunctionName:   InsertTitle
            Purpose:        插入标题
            Parameter:
                            1 workBook [QAxObject* &, IN]
                               工作簿指针

                            2 sheet [QAxObject* &, IN]
                               工作表指针

                            3 titleName [const QString&, IN]
                               标题名称

                            4 mode  [int, IN]
                               标题格式(训练成绩 or 学员成绩)
            Return:         void
            Remark:
        ***********************************************************/
	void InsertTitle(QAxObject* workBook, QAxObject* sheet, const QString& titleName, int mode);


    /*********************************************************
           FunctionName:   InsertTableTitle
           Purpose:        插入表头
           Parameter:
                           1 workBook [QAxObject* &, IN]
                              工作簿指针

                           2 sheet [QAxObject* &, IN]
                              工作表指针

                           3 mode  [int, IN]
                              表头格式(训练成绩 or 学员成绩)
           Return:         void
           Remark:
       ***********************************************************/
	void InsertTableTitle(QAxObject* workBook, QAxObject* sheet, int mode);


    /*********************************************************
            FunctionName:   InsertInfo
            Purpose:        插入记录
            Parameter:
                            1 workBook [QAxObject* &, IN]
                               工作簿指针

                            2 sheet [QAxObject* &, IN]
                               工作表指针

                            3~9 序号，学员名，搜索目标数，搜索时间，搜索距离，分值，排名

            Return:         void
            Remark:
        ***********************************************************/
	void InsertInfo(QAxObject* workBook, QAxObject* sheet, int count, const QString& userName, int serTargetCnt, 
		int serTime, int serDis, const QString&  grade, int rank);


    /*********************************************************
            FunctionName:   InsertInfo
            Purpose:        插入记录
            Parameter:
                            1 workBook [QAxObject* &, IN]
                               工作簿指针

                            2 sheet [QAxObject* &, IN]
                               工作表指针

                            3~7 序号，目标点，时间，距离，排名

            Return:         void
            Remark:
        ***********************************************************/
	void InsertInfo(QAxObject* workBook, QAxObject* sheet, int count, int targetPoint, int time, int dist, 
		int rank);



    /*********************************************************
        FunctionName:   CloseExcel
        Purpose:        关闭excel
        Parameter:
        Return:         void
        Remark:
                        关闭所有打开的工作簿、关闭excel进程
    ***********************************************************/
    void CloseExcel();

	//设置边框
	void SetBorder(QAxObject* workBook, QAxObject* sheet);

private:
    //判断工作簿是否打开
    bool IsOpened(const QString& workBookName);

    //判断表是否存在
    QAxObject* IsSheetExist(QAxObject* pSheets, const QString& sheetName);

    //获取一个工作簿对象
    QAxObject* GetOpenedWorkBook(const QString& workBookName);

    //合并单元格   1-起点单元格所在行 2-起点单元格所在行 3-终点单元格所在行 4-终点单元格所在列
    QAxObject* MergeCells(QAxObject* sheet, int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn);

    //获取表格中被使用的单元格的范围   1-起点单元格所在行 2-起点单元格所在行 3-终点单元格所在行 4-终点单元格所在列
    void GetUsageRange(QAxObject* sheet, int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn);

    //设置某单元格或区域格式   1-单元格对象  2-背景色 3-文字色 4-高度 5-宽度 6-文字大小 7-是否加粗
    void SetCell(QAxObject& cell, QColor& bgColor, QColor& txtColor, int height, int weight,int fontSize, bool isBold);

    //关闭所有工作簿
    void CloseAllWorkBook();


};

#endif // QEXCEL_H
