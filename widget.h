#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>
#include "qexcel.h"

namespace Ui {
class Widget;
}

class Widget : public QWidget
{
    Q_OBJECT

public:
    explicit Widget(QWidget *parent = 0);
    ~Widget();

private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

    void on_pushButton_3_clicked();

    void on_pushButton_4_clicked();

    void on_pushButton_5_clicked();

	void on_pushButton_6_clicked();

	void on_pushButton_7_clicked();

    void update();
private:
    Ui::Widget *ui;

    QExcel *m_qExcel;
};

#endif // WIDGET_H
