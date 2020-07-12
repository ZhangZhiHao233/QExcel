#include "widget.h"
#include "ui_widget.h"
#include <QString>
#include <QApplication>
#include<QTimer>

QAxObject *g_workBook;
QAxObject *g_sheet;

Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);
    m_qExcel = new QExcel();

    g_sheet = NULL;
    g_workBook = NULL;

    QTimer *timer = new QTimer(this);
    connect(timer, SIGNAL(timeout()), this, SLOT(update()));
    timer->start(1000);
}

Widget::~Widget()
{
    delete ui;
}

void Widget::on_pushButton_clicked()
{
  QString path = QApplication::applicationDirPath();
  g_workBook = m_qExcel->CreateWorkBook(g_sheet, path.append("/excel"), "cehsi");

}

void Widget::on_pushButton_2_clicked()
{
	m_qExcel->InsertTitle(g_workBook, g_sheet, "title", 1 );
}

void Widget::on_pushButton_3_clicked()
{
	m_qExcel->InsertTableTitle(g_workBook, g_sheet, 1);
}

void Widget::on_pushButton_4_clicked()
{
	m_qExcel->CloseWorkBook(g_workBook, "excel");
}

void Widget::on_pushButton_5_clicked()
{
	m_qExcel->CloseExcel();
}

void Widget::on_pushButton_6_clicked()
{
	m_qExcel->InsertInfo(g_workBook, g_sheet, 1, QStringLiteral("张三"), 22, 345, 1, QStringLiteral("第一"), 1);
}

void Widget::on_pushButton_7_clicked()
{
    m_qExcel->SetBorder(g_workBook, g_sheet);
}

void Widget::update()
{

}
