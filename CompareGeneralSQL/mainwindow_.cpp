#include "mainwindow.h"
#include "ui_mainwindow.h"



#define nullptr NULL

QString getCellValue(int row, int col, const QXlsx::Document &doc, bool boStrip = false)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        QVariant var = cell->readValue();
        asValue = var.toString();
    }
    else
    {
        asValue = "";
    }

    if (boStrip)
    {
      asValue = asValue.trimmed();
      asValue.replace("\r"," ").replace("\n", " ");
    }

    return asValue;

}

QString getCellValueValue(int row, int col, const QXlsx::Document &doc, bool boStrip = false)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        asValue = cell->value().toString();
    }
    else
    {
        asValue = "";
    }

    if (boStrip)
    {
      asValue = asValue.trimmed();
      asValue.replace("\r"," ").replace("\n", " ");
    }

    return asValue;

}


//ignore white spaces except one space
QString getCellValueWCIgnored(int row, int col, const QXlsx::Document &doc)
{
    QString asValue;
    Cell* cell = doc.cellAt(row, col);
    if(cell != nullptr)
    {
        QVariant var = cell->readValue();
        asValue = var.toString();
    }
    else
    {
        asValue = "";
    }

    asValue = asValue.simplified();
    asValue = asValue.toLower();
    asValue = asValue.trimmed();

    return asValue;

}

QString ignoreWhiteAndCase(QString asValue)
{
    asValue = asValue.simplified();
    asValue = asValue.toLower();

    return asValue;
}

QString ignoreWhiteCaseAndSpaces(QString asValue)
{
    asValue = asValue.simplified();
    asValue = asValue.toLower();
    asValue.replace(" ","");

    return asValue;
}

int CsvToExcel(QString fileNameCsv, QXlsx::Document &xlsxDocument)
{
    //7.6.2020 TODO .. include /r to text inside the quotes...until now not working...

    //Open file and convert it to QTextStrem

    QFile inputFile(fileNameCsv);
    if(!inputFile.open(QIODevice::ReadOnly))
    {
        return(knExitStatusCannotLoadInputCsvFiles);
    }
    QTextStream inputStream(&inputFile);


    //List of column lists
    QList <QStringList> lsCSVDoc;
    lsCSVDoc.clear();

    //Column list
    QStringList lstOneRowFields;
    lstOneRowFields.clear();

    QString asDelimiter = ";";
    QString asCurrentField = "";
    QString oneChar;

    //iterate over the characters
    bool boInQuotes = false;
    bool boQuotesInQuotes = false;
    bool boNewLine = false;
    lstOneRowFields.clear();
    while (!inputStream.atEnd())
    {
        boNewLine = false;
        oneChar = inputStream.read(1);
        if(oneChar == "\r") continue;

        //last character is " after quote section. Could be end of quote or escape for "
        if (boQuotesInQuotes)
        {
            if (oneChar == "\"")
            {
               boQuotesInQuotes = false;
               asCurrentField+=oneChar;
               continue;
            }
            else
            {
                 boQuotesInQuotes = false;
                 boInQuotes = false;
            }
        }
        else
        {

        }

        if (!boInQuotes && oneChar == "\"" )
        {
             boInQuotes = true;
        }
        else if (boInQuotes && oneChar== "\"" )
        {
            boQuotesInQuotes = true;
        }
        else if (!boInQuotes && oneChar == asDelimiter)
        {
            lstOneRowFields << asCurrentField;
            asCurrentField.clear();
        }
        else if (!boInQuotes && (oneChar == "\n"))
        {
            lstOneRowFields << asCurrentField;
            asCurrentField.clear();
            boNewLine = true;
            //new row
            lsCSVDoc << lstOneRowFields;
            lstOneRowFields.clear();
        }
        else
        {
            //do not add CR if it is outside of quoted text: if(!(!boInQuotes &&  oneChar == "\r"))  asCurrentField+=oneChar;
            asCurrentField+=oneChar;
        }

    }

    //if last LF in file is missing
    if(!boNewLine)
    {
      lstOneRowFields << asCurrentField;
      asCurrentField.clear();
      //new row
      lsCSVDoc << lstOneRowFields;
      lstOneRowFields.clear();
    }


    //write to Excel document
    int iCitacFields = 0;
    int iCitacRow = 0;
    Format text_num_format;
    text_num_format.setNumberFormat("@");
    foreach (QStringList lstRow, lsCSVDoc)
    {
        iCitacRow++;
        iCitacFields = 0;
        foreach (QString asItem, lstRow)
        {
            iCitacFields++;
            xlsxDocument.write(iCitacRow, iCitacFields, asItem, text_num_format);
            //qDebug() <<  iCitacRow  << QChar(64+iCitacFields) << ":" << asItem;
        }
    }

    return 0;
}




MainWindow::MainWindow(QStringList arguments, QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    setCentralWidget(ui->MainFrame);

    //display version
    setWindowTitle(QFileInfo( QCoreApplication::applicationFilePath() ).completeBaseName() + " V: " + QApplication::applicationVersion());
    ui->plainTextEdit_Statistic->setReadOnly(true);




    //other forms
    detailView = new DetailView(this);

    //init.
    newReqFileLastPath = ".//";
    oldReqFileLastPath = ".//";
    asCompareCondition = "";

    //ini files
    QString asIniFileName = QFileInfo( QCoreApplication::applicationFilePath() ).filePath().section(".",0,0)+".ini";
    iniSettings = new QSettings(asIniFileName, QSettings::IniFormat);
    QVariant VarTemp;



    VarTemp = iniSettings->value("paths/newReqFileLastPath");
    if (VarTemp.isValid())
    {
      newReqFileLastPath = VarTemp.toString();
    }
    else
    {
    }

    VarTemp = iniSettings->value("paths/oldReqFileLastPath");
    if (VarTemp.isValid())
    {
      oldReqFileLastPath = VarTemp.toString();
    }
    else
    {
    }

    VarTemp = iniSettings->value("extTools/WinMerge");
    if (VarTemp.isValid())
    {
      detailView->asWinMergePath = VarTemp.toString();
    }
    else
    {
      detailView->asWinMergePath = WIN_MERGE_LOCATION;
    }


    //set query table header
      QStringList tableHeader = QStringList() <<"Sel."
                                              <<"Requirement"
                                              <<"Short"
                                              <<"Status"
                                              <<"Parameter Name"
                                              <<"Old Value"
                                              <<"New Value"
                                              <<"User Note"
                                              << knOrigIDText
                                              ;

      ui->tableWidget_Changes->setColumnCount(tableHeader.count());
      ui->tableWidget_Changes->setHorizontalHeaderLabels(tableHeader);



      //get host and user name
      asUserAndHostName = "";
#if defined(Q_OS_WIN)
      char acUserName[100];
      DWORD sizeofUserName = sizeof(acUserName);
      if (GetUserNameA(acUserName, &sizeofUserName))asUserAndHostName = acUserName;
#endif
#if defined(Q_OS_LINUX)
      {
         QProcess process(this);
         process.setProgram("whoami");
         process.start();
         while (process.state() != QProcess::NotRunning) qApp->processEvents();
         asUserAndHostName = process.readAll();
         asUserAndHostName = asUserAndHostName.trimmed();
      }
#endif
      asUserAndHostName += "@" + QHostInfo::localHostName();
      //qDebug() << asUserAndHostName;

      boBatchProcessing = false;

      //CLI
      //queued connection - start after leaving constructor in case of batch processing
      connect(this,SIGNAL(startBatchProcessing(int)),
              SLOT(batchProcessing(int)),
              Qt::QueuedConnection);


      //process command line arguments
      comLineArgList = arguments;
      ////remove exec name
      if(comLineArgList.size() > 0) comLineArgList.removeFirst();

      if(comLineArgList.size() >= 2)
      {
        ui->lineEdit_NewReq->setText(comLineArgList.at(0).trimmed());
        comLineArgList.removeFirst();
        ui->lineEdit_OldReq->setText(comLineArgList.at(0).trimmed());
        comLineArgList.removeFirst();


        if(comLineArgList.contains("-a")) ui->cbOnlyFuncReq->setChecked(false);
        else                              ui->cbOnlyFuncReq->setChecked(true);

        if(comLineArgList.contains("-w")) ui->cb_IngnoreWaC->setChecked(true);
        else                              ui->cb_IngnoreWaC->setChecked(false);

        if(comLineArgList.contains("-p")) ui->cb_IgnoreProject->setChecked(true);
        else                              ui->cb_IgnoreProject->setChecked(false);

        if(comLineArgList.contains("-l")) ui->cb_IgnoreLinks->setChecked(true);
        else                              ui->cb_IgnoreLinks->setChecked(false);


        //check for batch processing (without window)
        boBatchProcessing = false;
        iExitCode = 0;
        if(comLineArgList.contains("-b"))
        {
           //do not show main window, process without it
           boBatchProcessing = true;
           startBatchProcessing(knBatchProcessingID);
           //qDebug() << "BATCH PROCESSING STARTED, SHOULD BE FIRST";

        }

     }

     //Prepare SQL Lite database
     db_Req = QSqlDatabase::addDatabase("QSQLITE","sqlite_connection");
     db_Req.setDatabaseName(":memory:");
     bool boGenResult = db_Req.open();
     Q_UNUSED(boGenResult);
     //qDebug() << "opendb" << boGenResult;

     modelOld =  new QSqlTableModel(this, db_Req);
     modelNew =  new QSqlTableModel(this, db_Req);
     modelUserQuery = new QSqlQueryModel();


     EmptyChangesTable();

     //set visible "compare" page in result tab
     ui->tabWidget_Results->setCurrentWidget(ui->pageCompare);


     //init statistic Variables
     iStatisticOldFile      = -1;
     iStatisticNewFile      = -1;
     iStatisticChangedDiff  = -1;
     iStatisticChangedReq   = -1;
     iStatisticNewReq       = -1;
     iStatisticMissingReq       = -1;


}


int MainWindow::batchProcessing(int iID)
{
  iExitCode = 0;
  if (iID != knBatchProcessingID) iExitCode = knExitStatusBadSignal;
  if(!iExitCode)  this->on_btnCompare_clicked();
  if(!iExitCode)  this->on_btnWrite_clicked();

  QCoreApplication::exit(iExitCode);

  return(iExitCode);


}

void MainWindow::EmptyChangesTable()
{
    //delete old if exist
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           delete ui->tableWidget_Changes->item(irow, icol);
           ui->tableWidget_Changes->setItem(irow, icol, NULL);
         }

       }
     }
     ui->tableWidget_Changes->setRowCount(0);
     ui->btnWrite->setEnabled(false);

    QSqlQuery dbQuery(db_Req);
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tOLD`");
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tNEW`");
    ui->listWidget_oldColumns->clear();
    ui->listWidget_newColumns->clear();
    ui->listWidget_Commands->clear();
    modelUserQuery->clear();
    ui->tableView_Query->setModel(modelUserQuery);

    //clear statistic list
    ui->plainTextEdit_Statistic->clear();
    ui->plainTextEdit_Statistic->appendPlainText("STATISTIC WILL");
    ui->plainTextEdit_Statistic->appendPlainText("BE UPDATED");
    ui->plainTextEdit_Statistic->appendPlainText("AFTER WRITE");


}


void MainWindow::setChangesTableNotEditable()
{
    //set table not editable, not selectable - exeptions is user comment column
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {

           if (icol != col_infotable_user_comment) //user comment of course editable...
           {


             QTableWidgetItem *item = ui->tableWidget_Changes->item(irow, icol);
             if(item== nullptr)
             {
                item = new QTableWidgetItem();
                ui->tableWidget_Changes->setItem(irow,icol, item);
             }
             item->setFlags(item->flags() & ~(Qt::ItemIsEditable) & ~(Qt::ItemIsSelectable));
          }
       }
     }
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btnWrite_clicked()
{

    //write result
    QXlsx::Document reportDoc;

    //add document properties
    reportDoc.setDocumentProperty("creator", asUserAndHostName);
    reportDoc.setDocumentProperty("description", "Requirements Comparsion");


    reportDoc.addSheet("TABLE");

    int iReportCurrentRow = kn_ReportFirstInfoRow;
    //info to report file
    reportDoc.write(iReportCurrentRow++, 1, "generated by: " +
                                          QFileInfo( QCoreApplication::applicationName()).fileName() +
                                          " V:" + QApplication::applicationVersion() +
                                          ", "  + asUserAndHostName);
    reportDoc.write(iReportCurrentRow++, 1, "generated on: " + QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss"));
    reportDoc.write(iReportCurrentRow++, 1, "new file: " + fileName_NewReq);
    reportDoc.write(iReportCurrentRow++, 1, "old file: " + fileName_OldReq);
    reportDoc.write(iReportCurrentRow++, 1, asCompareCondition);
    int iRowWhereToWriteNumberOfReqWritten = iReportCurrentRow++;

// one blank line
     iReportCurrentRow++;

// headers
     Format font_bold;
     font_bold.setFontBold(true);
     reportDoc.write(iReportCurrentRow,     1, "Requirement", font_bold);       reportDoc.setColumnWidth(1, 23);
     reportDoc.write(iReportCurrentRow,     2, "Short", font_bold);             reportDoc.setColumnWidth(2, 8);
     reportDoc.write(iReportCurrentRow,     3, "Status", font_bold);            reportDoc.setColumnWidth(3, 16);
     reportDoc.write(iReportCurrentRow,     4, "Parameter Name", font_bold);    reportDoc.setColumnWidth(4, 25);
     reportDoc.write(iReportCurrentRow,     5, "Old Value", font_bold);         reportDoc.setColumnWidth(5, 75);
     reportDoc.write(iReportCurrentRow,     6, "New Value", font_bold);         reportDoc.setColumnWidth(6, 75);
     reportDoc.write(iReportCurrentRow,     7, "User Notes", font_bold);        reportDoc.setColumnWidth(7, 50);
     reportDoc.write(iReportCurrentRow,     8, knOrigIDText, font_bold);        reportDoc.setColumnWidth(8, 23);


     //new row
     iReportCurrentRow++;
     int iReqWritten = 0;

     QSet<QString> setStatisticReq;
     setStatisticReq.clear();
     iStatisticNewReq          = 0;
     iStatisticMissingReq      = 0;
     iStatisticChangedDiff     = 0;
     //copy from table to excel sheet "TABLE"
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {
       //it is selected by user (check box)
       QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
       if(checkBox != nullptr)
       {
           if(!(checkBox->isChecked())) continue; //not selected, next row
       }
       iReqWritten++;

       //statistic
       if(ui->tableWidget_Changes->item(irow, col_infotable_status)->text().toLower().contains("changed"))
       {
         iStatisticChangedDiff++;
         //eliminate same requirement (with differencies in more column)
         setStatisticReq << ui->tableWidget_Changes->item(irow, col_infotable_requirement)->text();
       }

       if(ui->tableWidget_Changes->item(irow, col_infotable_status)->text().toLower().contains("missing"))
       {
         iStatisticMissingReq++;
       }

       if(ui->tableWidget_Changes->item(irow, col_infotable_status)->text().toLower().contains("new"))
       {
         iStatisticNewReq++;
       }

       //write to excel
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         int iExcelCol = icol;
         if(iExcelCol < 1) continue;
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           Format cellFormat;
           QColor cellColor = ui->tableWidget_Changes->item(irow, icol)->background().color();
           if((cellColor.isValid()) && (cellColor != Qt::black))
           {
               //qDebug() << cellColor;
               cellFormat.setPatternBackgroundColor(cellColor);
           }


           QString asTableContent = ui->tableWidget_Changes->item(irow, icol)->text();
           reportDoc.write(iReportCurrentRow, iExcelCol, asTableContent, cellFormat);                   
         }

       }
       iReportCurrentRow++;
     }

     iStatisticChangedReq = setStatisticReq.count();
     //fill statistic list
     ui->plainTextEdit_Statistic->clear();
     ui->plainTextEdit_Statistic->appendPlainText(QString("in Old File: %1").arg(iStatisticOldFile));
     ui->plainTextEdit_Statistic->appendPlainText(QString("in New File: %1").arg(iStatisticNewFile));
     ui->plainTextEdit_Statistic->appendPlainText(QString("Differencies: %1").arg(iStatisticChangedDiff));
     ui->plainTextEdit_Statistic->appendPlainText(QString("Changed Req: %1").arg(iStatisticChangedReq));
     ui->plainTextEdit_Statistic->appendPlainText(QString("New Req: %1").arg(iStatisticNewReq));
     ui->plainTextEdit_Statistic->appendPlainText(QString("Missing Req: %1").arg(iStatisticMissingReq));

     qDebug() << ui->plainTextEdit_Statistic->toPlainText().replace("\r","").replace("\n","::");

     QString asReqWritten = "";
     if (iReqWritten < ui->tableWidget_Changes->rowCount())
     {
         asReqWritten = QString("Only %1 table row of %2 written").arg(iReqWritten).arg(ui->tableWidget_Changes->rowCount());
     }
     else
     {
         asReqWritten = QString("All %1 table row written").arg(ui->tableWidget_Changes->rowCount());
     }

     reportDoc.write(iRowWhereToWriteNumberOfReqWritten, 1,
                     asReqWritten + "::"+ui->plainTextEdit_Statistic->toPlainText().replace("\r","").replace("\n","::"));






     //write excel sheet "EASY TO READ"
     reportDoc.addSheet("EASY TO READ");

     reportDoc.setColumnWidth(2, 23);
     reportDoc.setColumnWidth(3, 8);
     reportDoc.setColumnWidth(4, 16);
     reportDoc.setColumnWidth(5, 25);

     int iWriteToExcelRow = -1;
     int iWriteToExcelCol = -1;

     iWriteToExcelRow = 1;
     for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
     {

       //it is selected by user (check box)
       QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
       if(checkBox != nullptr)
       {
          if(!(checkBox->isChecked())) continue; //not selected, next row
       }


       bool boExtraRow = false;
       for (int icol = 0; icol < ui->tableWidget_Changes->columnCount(); ++icol)
       {
         if(icol == col_infotable_origID) continue;
         if(ui->tableWidget_Changes->item(irow, icol) != NULL)
         {
           Format cellFormat;
           QColor cellColor = ui->tableWidget_Changes->item(irow, icol)->background().color();
           if((cellColor.isValid()) && (cellColor != Qt::black))
           {
               //qDebug() << cellColor;
               cellFormat.setPatternBackgroundColor(cellColor);
           }

           if(icol == col_infotable_requirement)
           {
               cellFormat.setFontBold(true);
           }

           QString asTableContent = ui->tableWidget_Changes->item(irow, icol)->text();
           asTableContent = asTableContent.trimmed();
           if(asTableContent.isEmpty() || asTableContent.isNull())
           {
               //boExtraRow = true;
               continue;
           }

           iWriteToExcelCol = icol+1;
           if(icol == col_infotable_old_value)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "OLD: "+asTableContent;
             boExtraRow = true;
           }

           if(icol == col_infotable_new_value)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "NEW: "+asTableContent;
             boExtraRow = true;
           }

           if(icol == col_infotable_user_comment)
           {
             iWriteToExcelRow++;
             iWriteToExcelCol = 1;
             asTableContent = "USER: "+asTableContent;
             boExtraRow = true;
           }
           reportDoc.write(iWriteToExcelRow, iWriteToExcelCol, asTableContent, cellFormat);

           //write origin ID behind parameter name
           if(icol == col_infotable_par_name)
           {
              if(ui->tableWidget_Changes->item(irow, col_infotable_origID) != NULL)
              {
                 asTableContent = ui->tableWidget_Changes->item(irow, col_infotable_origID)->text();
                 asTableContent = asTableContent.trimmed();
                 if(!asTableContent.isEmpty()) asTableContent = "(" +asTableContent + ")";
                 reportDoc.write(iWriteToExcelRow, iWriteToExcelCol+1, asTableContent, cellFormat);
              }
           }



         }

       }
       iWriteToExcelRow++;
       if(boExtraRow) iWriteToExcelRow++;

     }

    reportDoc.selectSheet("TABLE");
    if(true)
    {
       if(!reportDoc.saveAs(fileName_Report))
       {

           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Error, not opened?", QMessageBox::Ok);
           }
           else
           {
              qCritical() << "Error: "<< "Problem write to report file (opened?)";
              iExitCode = knExitStatusReporFileOpened;
           }

       }
       else
       {

           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Written", fileName_Report+"\r\n\r\n"+asReqWritten, QMessageBox::Ok);
           }
           else
           {
              qWarning() << "Success: " << asReqWritten;
           }

       }

    }
    else
    {
        if (!boBatchProcessing)
        {
            QMessageBox::information(this, "Problem", "No result doc", QMessageBox::Ok);
        }
        else
        {
           qCritical() << "Error: " << "No result doc";
           iExitCode = knExitStatusNoResultDoc;
        }

    }
}

void MainWindow::on_btnNewReq_clicked()
{
    EmptyChangesTable();
    // local only, it is use later from edit window
    QString fileName_NewReq = QFileDialog::getOpenFileName
    (
      this,
      "Open New Requirements",
      newReqFileLastPath,
      "Excel Files (*.xlsx; *.csv)"
    );

    if(!fileName_NewReq.isEmpty() && !fileName_NewReq.isNull())
    {
        ui->lineEdit_NewReq->setText(fileName_NewReq);
        newReqFileLastPath = QFileInfo(fileName_NewReq).path();
    }
    else
    {
        ui->lineEdit_NewReq->clear();
    }


}


void MainWindow::on_lineEdit_NewReq_textEdited(const QString &arg1)
{
    Q_UNUSED(arg1)
    //new files - table is not valid for them
    EmptyChangesTable();
}

void MainWindow::on_btnOldReq_clicked()
{
    EmptyChangesTable();
    // local only, it is use later from edit window
    QString fileName_OldReq = QFileDialog::getOpenFileName
    (
      this,
      "Open Old Requirements",
      oldReqFileLastPath,
      "Excel Files (*.xlsx; *.csv)"
    );

    if(!fileName_OldReq.isEmpty() && !fileName_OldReq.isNull())
    {
        ui->lineEdit_OldReq->setText(fileName_OldReq);
        oldReqFileLastPath = QFileInfo(fileName_OldReq).path();
    }
    else
    {
        ui->lineEdit_OldReq->clear();
    }
}

void MainWindow::on_lineEdit_OldReq_textEdited(const QString &arg1)
{
      Q_UNUSED(arg1)
    //new files - table is not valid for them
      EmptyChangesTable();
}

void MainWindow::on_btnCompare_clicked()
{
    //set visible "compare" page in result tab
    ui->tabWidget_Results->setCurrentWidget(ui->pageCompare);

    fileName_NewReq = ui->lineEdit_NewReq->text();
    fileName_OldReq = ui->lineEdit_OldReq->text();

    //check validity of filenames
    if(
            fileName_NewReq.isEmpty() ||
            fileName_NewReq.isNull()  ||
            fileName_OldReq.isEmpty() ||
            fileName_OldReq.isNull()
        )
    {
        if (!boBatchProcessing)
        {
            QMessageBox::information(this, "No filenames", "Select Files", QMessageBox::Ok);
        }
        else
        {
           qCritical() << "Error: " << "Invalid input filenames";
           iExitCode = knExitStatusInvalidInputFilenames;
        }
        return;
    }






    //Load the new document or convert it from csv file
    qDebug() << __LINE__;
    QXlsx::Document  newReqDoc(fileName_NewReq);
    qDebug() << __LINE__;

    if(QFileInfo(fileName_NewReq).suffix().toLower() == "csv")
    {
       int iResult = CsvToExcel(fileName_NewReq, newReqDoc);
       if(iResult)
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load new csv file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Problem to load new csv files";
               iExitCode = iResult;
           }
           return;
       }
    }
    else
    {
       if (!newReqDoc.load())
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load new file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Can not load new input file";
               iExitCode = knExitStatusCannotLoadInputFiles;
           }

           return;
       }
    }

    //newReqDoc.saveAs("pn.xlsx");  //DEBUG CODE

    //Load the old document or convert it from csv file
    QXlsx::Document  oldReqDoc(fileName_OldReq);

    if(QFileInfo(fileName_OldReq).suffix().toLower() == "csv")
    {
       int iResult = CsvToExcel(fileName_OldReq, oldReqDoc);
       if(iResult)
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load old csv file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Problem to load old csv files";
               iExitCode = iResult;
           }
           return;
       }
    }
    else
    {
       if (!oldReqDoc.load())
       {
           if (!boBatchProcessing)
           {
               QMessageBox::information(this, "Problem", "Problem to load old file", QMessageBox::Ok);
           }
           else
           {
               qCritical() << "Error: " << "Can not load old input files";
               iExitCode = knExitStatusCannotLoadInputFiles;
           }

           return;
       }
    }

    //oldReqDoc.saveAs("po.xlsx");  //DEBUG CODE





    //save paths to ini file
    iniSettings->setValue("paths/newReqFileLastPath", newReqFileLastPath);
    iniSettings->setValue("paths/oldReqFileLastPath", oldReqFileLastPath);


    //prepare report file
    fileName_Report =   QFileInfo(fileName_NewReq).path() + "/" +
                        QFileInfo(fileName_NewReq).completeBaseName() +
                        "_to_" +
                        QFileInfo(fileName_OldReq).completeBaseName()+
                        ".xlsx";

    fileName_OldReqCor = QFileInfo(fileName_OldReq).path() + "/" +
                         QFileInfo(fileName_OldReq).completeBaseName()+ 
                         "_cor"
                         ".xlsx";                         

    //qDebug() << fileName_NewReq << fileName_OldReq << fileName_Report;


    int newLastRow = newReqDoc.dimension().lastRow();
    int newLastCol = newReqDoc.dimension().lastColumn();
    int oldLastRow = oldReqDoc.dimension().lastRow();
    int oldLastCol = oldReqDoc.dimension().lastColumn();

    int maxLastRow = (newLastRow > oldLastRow) ? newLastRow : oldLastRow;
    int maxLastCol = (newLastCol > oldLastCol) ? newLastCol : oldLastCol;

    Q_UNUSED(maxLastRow)
    Q_UNUSED(maxLastCol)

    //qDebug() <<newLastRow << newLastCol << oldLastRow << oldLastCol << maxLastRow << maxLastCol;




    //Delete changes table
    EmptyChangesTable();

    //disable sorting
    ui->tableWidget_Changes->setSortingEnabled(false);
    //close detail
    detailView->close();


    //start compare demo

    /*
    for(int row =1; row <= maxLastRow; row++)
    {
        for(int col=1; col <= maxLastCol; col++)
        {
           //qDebug() << getCellValue(row, col, newReqDoc)  << getCellValue(row, col, oldReqDoc);
           reportDoc->write(row, col, QVariant(getCellValue(row, col, newReqDoc)+ "/" + getCellValue(row, col, oldReqDoc)));
        }
    }
    */


    //start compare real
    //find position of "Object Type" Column and add headers to list

    QStringList lstNewHeaders;
    lstNewHeaders.clear();
    QMap<QString, int> mapNewHeaders;
    mapNewHeaders.clear();

    for(int col= kn_ReqIDCol; col <= newLastCol; col++)
    {
      QString asTemp = getCellValue(kn_HeaderRow, col, newReqDoc, true);
      lstNewHeaders << asTemp;
      mapNewHeaders[asTemp] = col;
    }

        //Test Columns
        foreach (QString asTemp, lstNewHeaders)
        {
          qDebug() << "new: " << asTemp;
        }


    QStringList lstOldHeaders;
    lstOldHeaders.clear();
    QMap<QString, int> mapOldHeaders;
    mapOldHeaders.clear();

    for(int col= kn_ReqIDCol; col <= oldLastCol; col++)
    {
        QString asTemp = getCellValue(kn_HeaderRow, col, oldReqDoc, true);
        lstOldHeaders << asTemp;
        mapOldHeaders[asTemp] = col;
    }

    //Test Columns
    foreach (QString asTemp, lstOldHeaders)
    {
      qDebug() << "old: " << asTemp;
    }





    //TODO exit when Object Type not found


    QStringList lstNewReqIDs;
    lstNewReqIDs.clear();
    QStringList lstOldReqIDs;
    lstOldReqIDs.clear();

    //Create database table OLD
    QSqlQuery dbQuery(db_Req);

    ui->tableView_Old->setSortingEnabled(false);
    ui->listWidget_oldColumns->clear();
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tOLD`");
    boGenResult = dbQuery.exec("CREATE TABLE `tOLD` (rowid INTEGER PRIMARY KEY)");
    //qDebug() <<"Create" << boGenResult;
    foreach (QString asColumnName, lstOldHeaders)
    {
       boGenResult = dbQuery.exec("ALTER TABLE `tOLD` ADD \""+asColumnName+"\" VARCHAR" );
       //qDebug() <<"Add" << asColumnName << boGenResult;
       ui->listWidget_oldColumns->addItem(asColumnName);
    }

    //insert data to OLD
    iStatisticOldFile = oldLastRow - kn_FistDataRow + 1;
    bool boInsertedOK = true;
    for (int oldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)
    {
        QString asColumns = "";
        QString asValues  = "";
        foreach (QString asColumnName, lstOldHeaders)
        {
          asColumns += "`"+asColumnName+"`";
          asColumns +=", ";
          QString asValue = getCellValueValue(oldRow, mapOldHeaders[asColumnName], oldReqDoc, true);
          asValue.replace("'","''");
          asValues += "'"+asValue+"'";
          asValues +=", ";
        }
        //remove last comma..
        asColumns = asColumns.left(asColumns.lastIndexOf(","));
        asValues = asValues.left(asValues.lastIndexOf(","));

        QString asQueryString = "INSERT INTO `tOLD` (" +asColumns+ ") values(" + asValues + ")";
        //qDebug() << "old: " << asQueryString;
        boGenResult = dbQuery.exec(asQueryString);
        if(!boGenResult) boInsertedOK = false;
        //qDebug() <<"Insert" << boGenResult;

    }

    //report problem OLD
    if(!boInsertedOK)
    {
      if (!boBatchProcessing)
      {
         QMessageBox::information(this, "Error", "Problem to insert OLD to db table", QMessageBox::Ok);
      }
      else
      {
        qWarning() << "Error " <<  "Problem to insert OLD to db table";
        //TODO exit code
      }
    }

    //show in table OLD
    modelOld->clear();
    modelOld->setTable("tOLD");
    modelOld->select();
    qDebug() << "modelOld->rowCount()" << modelOld->rowCount();

    ui->tableView_Old->setModel(modelOld);

//    //adapt column with OLD
//    ui->tableView_Old->resizeColumnsToContents();
//    for(int i=0; i<ui->tableView_Old->model()->columnCount(); i++)
//    {
//     if(ui->tableView_Old->columnWidth(i) > knMAX_VIEW_COLUMN_WIDTH)
//       ui->tableView_Old->setColumnWidth(i, knMAX_VIEW_COLUMN_WIDTH);
//    }
//    ui->tableView_Old->setSortingEnabled(true);



    //Create database table NEW
    ui->tableView_New->setSortingEnabled(false);
    ui->listWidget_newColumns->clear();
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tNEW`");
    boGenResult = dbQuery.exec("CREATE TABLE `tNEW` (rowid INTEGER PRIMARY KEY)");
    //qDebug() <<"Create" << boGenResult;
    foreach (QString asColumnName, lstNewHeaders)
    {
       boGenResult = dbQuery.exec("ALTER TABLE `tNEW` ADD \""+asColumnName+"\" VARCHAR" );
       //qDebug() <<"Add" << asColumnName << boGenResult;
       ui->listWidget_newColumns->addItem(asColumnName);
    }


    //insert data NEW
    iStatisticNewFile = newLastRow - kn_FistDataRow + 1;
    boInsertedOK = true;
    for (int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
    {
      QString asColumns = "";
      QString asValues  = "";
      foreach (QString asColumnName, lstNewHeaders)
      {
        asColumns += "`"+asColumnName+"`";
        asColumns +=", ";
        QString asValue = getCellValueValue(newRow, mapNewHeaders[asColumnName], newReqDoc, true);
        asValue.replace("'","''");
        asValues += "'"+asValue+"'";
        asValues +=", ";
      }
      //remove last comma..
      asColumns = asColumns.left(asColumns.lastIndexOf(","));
      asValues = asValues.left(asValues.lastIndexOf(","));

      QString asQueryString = "INSERT INTO `tNEW` (" +asColumns+ ") values(" + asValues + ")";
      //qDebug() << "new: " << asQueryString;
      boGenResult = dbQuery.exec(asQueryString);
      if(!boGenResult) boInsertedOK = false;
      //qDebug() <<"Insert" << boGenResult;

    }

    //report problem NEW
    if(!boInsertedOK)
    {
      if (!boBatchProcessing)
      {
         QMessageBox::information(this, "Error", "Problem to insert NEW to db table", QMessageBox::Ok);
      }
      else
      {
        qWarning() << "Error " <<  "Problem to insert NEW to db table";
        //TODO exit code
      }
    }

    //show in table NEW
    modelNew->clear();
    modelNew->setTable("tNEW");
    modelNew->select();
    qDebug() << "modelNew->rowCount()" << modelNew->rowCount();
    ui->tableView_New->setModel(modelNew);

//    ui->tableView_New->resizeColumnsToContents();
//    for(int i=0; i<ui->tableView_New->model()->columnCount(); i++)
//    {
//      if(ui->tableView_New->columnWidth(i) > knMAX_VIEW_COLUMN_WIDTH)
//        ui->tableView_New->setColumnWidth(i, knMAX_VIEW_COLUMN_WIDTH);
//    }
//    ui->tableView_New->setSortingEnabled(true);


    //insert key word to list for SQL query
    ui->listWidget_Commands->clear();
    ui->listWidget_Commands->addItem("SELECT");
    ui->listWidget_Commands->addItem("*");
    ui->listWidget_Commands->addItem("FROM");
    ui->listWidget_Commands->addItem("`tNEW`");
    ui->listWidget_Commands->addItem("`tOLD`");
    ui->listWidget_Commands->addItem("WHERE");
    ui->listWidget_Commands->addItem("JOIN");
    ui->listWidget_Commands->addItem("LEFT JOIN");
    ui->listWidget_Commands->addItem("INNER JOIN");
    ui->listWidget_Commands->addItem("ON");
    ui->listWidget_Commands->addItem("'");
    ui->listWidget_Commands->addItem("`");
    ui->listWidget_Commands->addItem("=");
    ui->listWidget_Commands->addItem("ORDER");
    ui->listWidget_Commands->addItem("BY");
    ui->listWidget_Commands->addItem("ASC");
    ui->listWidget_Commands->addItem("DESC");
    ui->listWidget_Commands->addItem("AS");



}

void MainWindow::on_tableWidget_Changes_clicked(const QModelIndex &index)
{
    //qDebug() << index.row() << index.column();


    QString asOldText = "";
    QString asNewText = "";
    QString asReqIDandParam ="???";


    //texts
    if(ui->tableWidget_Changes->item(index.row(), col_infotable_old_value) != nullptr)
      asOldText = ui->tableWidget_Changes->item(index.row(), col_infotable_old_value)->text();
    if( ui->tableWidget_Changes->item(index.row(), col_infotable_new_value) != nullptr)
      asNewText = ui->tableWidget_Changes->item(index.row(), col_infotable_new_value)->text();
    detailView->setTexts(asOldText, asNewText);

    //title  = req ID
    if(ui->tableWidget_Changes->item(index.row(), col_infotable_requirement) != nullptr)
      asReqIDandParam = ui->tableWidget_Changes->item(index.row(), col_infotable_requirement)->text();
    asReqIDandParam += "/";
    if(ui->tableWidget_Changes->item(index.row(), col_infotable_par_name) != nullptr)
      asReqIDandParam += ui->tableWidget_Changes->item(index.row(), col_infotable_par_name)->text();


    detailView->setReqID(asReqIDandParam);

    if(index.column() == col_infotable_check_box) return; //select field
    if(index.column() == col_infotable_user_comment) return; //user notes field

    detailView->show();


}

void MainWindow::on_btn_Debug_clicked()
{
    //if(detailView != nullptr) detailView->exec();
    detailView->show();
    detailView->setTexts("My Old", "My New");
}





void MainWindow::on_btn_CheckAll_clicked()
{
  bool static boNowCheckAll = true;
  if(boNowCheckAll) checkAll(true);
  else              checkAll(false);
  boNowCheckAll = !boNowCheckAll;
}


void MainWindow::checkAll(bool boCheck)
{
    for (int irow = 0; irow < ui->tableWidget_Changes->rowCount(); ++irow)
    {
        QCheckBox *checkBox = dynamic_cast<QCheckBox *> (ui->tableWidget_Changes->cellWidget(irow, col_infotable_check_box));
        if(checkBox != nullptr)  checkBox->setChecked(boCheck);
    }

}

void MainWindow::on_btn_SortReq_clicked()
{
   bool static boSortAsc = true;
   if(boSortAsc)
     ui->tableWidget_Changes->sortByColumn(col_infotable_requirement, Qt::AscendingOrder);
   else
     ui->tableWidget_Changes->sortByColumn(col_infotable_requirement, Qt::DescendingOrder);


   boSortAsc = !boSortAsc;
}

void MainWindow::on_btn_ExecQuery_clicked()
{

    ui->tableView_Query->setSortingEnabled(false);

    QString asQueryText = ui->lineEdit_Query->text();
    QSqlQuery userQuery(db_Req);
    bool boResult = userQuery.exec(asQueryText);
    if(!boResult)
    {
      QMessageBox::information(this, "Result", userQuery.lastError().text(), QMessageBox::Ok);
      modelUserQuery->clear();
      ui->tableView_Query->setModel(modelUserQuery);
    }
    else
    {
      modelUserQuery->setQuery(userQuery);
      ui->tableView_Query->setModel(modelUserQuery);
      ui->tableView_Query->resizeColumnsToContents();
      for(int i=0; i<ui->tableView_Query->model()->columnCount(); i++)
      {
        if(ui->tableView_Query->columnWidth(i) > knMAX_VIEW_COLUMN_WIDTH)
          ui->tableView_Query->setColumnWidth(i, knMAX_VIEW_COLUMN_WIDTH);
      }
      //ui->tableView_Query->setSortingEnabled(true);  //not working - you have to subclass and
    }
}

void MainWindow::on_listWidget_Commands_itemClicked(QListWidgetItem *item)
{
    addTextToQueryLine(item->text());
}

void MainWindow::on_listWidget_oldColumns_itemClicked(QListWidgetItem *item)
{
    QString asToAdd = "`"+ item->text() + "`";
    if(ui->cb_AddTableNameToColName->isChecked())
        asToAdd = "`tOLD`."+ asToAdd;
    addTextToQueryLine(asToAdd);
}

void MainWindow::on_listWidget_newColumns_itemClicked(QListWidgetItem *item)
{
    QString asToAdd = "`"+ item->text() + "`";
    if(ui->cb_AddTableNameToColName->isChecked())
        asToAdd = "`tNEW`."+ asToAdd;
    addTextToQueryLine(asToAdd);
}

void MainWindow::addTextToQueryLine(QString asToAdd)
{
  ui->lineEdit_Query->insert(asToAdd + " ");
}

void MainWindow::on_btn_ClearQueryLine_clicked()
{
   ui->lineEdit_Query->clear();
}
