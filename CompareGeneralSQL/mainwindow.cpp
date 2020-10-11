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


QString ignoreWhiteAndCase(QString asValue)
{
    asValue = asValue.simplified();
    asValue = asValue.toUpper();

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




    //other forms
    manageStoredQueries = new ManageQueris(this);

    //init.
    newReqFileLastPath = ".//";
    oldReqFileLastPath = ".//";
    asCompareCondition = "";

    //ini files
    QString asIniFileName = QFileInfo( QCoreApplication::applicationFilePath() ).filePath().section(".",0,0)+".ini";
    iniSettings = new QSettings(asIniFileName, QSettings::IniFormat);
    //in other forms:
    manageStoredQueries->currentSettings = iniSettings;

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
    }
    else
    {
    }





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



        if(comLineArgList.contains("-w")) ui->cb_IngnoreWaC->setChecked(true);
        else                              ui->cb_IngnoreWaC->setChecked(false);




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
     db_Req.close();
     db_Req.open();
     bool boGenResult = db_Req.open();
     Q_UNUSED(boGenResult);
     //qDebug() << "opendb" << boGenResult;

     modelOld =  new QSqlTableModel(this, db_Req);
     modelNew =  new QSqlTableModel(this, db_Req);
     modelUserQuery = new QSqlQueryModel();


     EmptyChangesTable();

     //set visible "compare" page in result tab
     ui->tabWidget_Results->setCurrentWidget(ui->pageQuery);




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

     ui->btnWrite->setEnabled(false);

    QSqlQuery dbQuery(db_Req);
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tOLD`");
    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tNEW`");
    ui->listWidget_oldColumns->clear();
    ui->listWidget_newColumns->clear();
    ui->listWidget_Commands->clear();
    modelUserQuery->clear();
    ui->tableView_Query->setModel(modelUserQuery);



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


    reportDoc.addSheet("QEURY TABLE");

    int iReportCurrentRow = kn_ReportFirstInfoRow;
    //info to report file
    reportDoc.write(iReportCurrentRow++, 1, "generated by: " +
                                          QFileInfo( QCoreApplication::applicationName()).fileName() +
                                          " V:" + QApplication::applicationVersion() +
                                          ", "  + asUserAndHostName);
    reportDoc.write(iReportCurrentRow++, 1, "generated on: " + QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss"));
    reportDoc.write(iReportCurrentRow++, 1, "new file: " + fileName_NewReq  + ", sheet: " + asNewSheetName);
    reportDoc.write(iReportCurrentRow++, 1, "old file: " + fileName_OldReq  + ", sheet: " + asOldSheetName);
    reportDoc.write(iReportCurrentRow++, 1, asCompareCondition);
    reportDoc.write(iReportCurrentRow++, 1, "query executed: " + asQueryTextToReport);

// one blank line
     iReportCurrentRow++;




// write headers to excel
     Format font_bold;
     font_bold.setFontBold(true);

     for (int iIndexCol = 0; iIndexCol < modelUserQuery->columnCount(); ++iIndexCol)
     {
        QString asHeaderName = modelUserQuery->headerData(iIndexCol, Qt::Horizontal).toString();
        reportDoc.write(iReportCurrentRow, iIndexCol+1, asHeaderName,  font_bold);
     }

//write data to excel
     for (int iIndexRow = 0; iIndexRow < modelUserQuery->rowCount(); ++iIndexRow)
     {
         iReportCurrentRow++;
         QString asRowText = "";
         for (int iIndexCol = 0; iIndexCol < modelUserQuery->columnCount(); ++iIndexCol)
         {
            QString asItemText = modelUserQuery->data(modelUserQuery->index(iIndexRow, iIndexCol)).toString();
            asRowText += asItemText;
            reportDoc.write(iReportCurrentRow, iIndexCol+1, asItemText);

            if((iIndexCol + 1) <  modelUserQuery->columnCount()) asRowText +=";";
         }
         //qDebug() << asRowText;
     }

//save excel to file
     QString fileName_Report_complet = fileName_Report + ui->lineEdit_ReportSuffix->text() + ".xlsx";
     if(!reportDoc.saveAs(fileName_Report_complet))
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
             QMessageBox::information(this, "Written", fileName_Report_complet +"\r\n\r\n", QMessageBox::Ok);
         }
         else
         {
            qWarning() << "Success ";
         }

     }

}

void MainWindow::on_btnNewReq_clicked()
{
    EmptyChangesTable();
    ui->comboBox_SheetsNew->clear();
    //delete sheet names
    asNewSheetName.clear();




    // local only, it is use later from edit window
    QString fileName_NewReq = QFileDialog::getOpenFileName
    (
      this,
      "Open New Requirements",
      newReqFileLastPath,
      "Excel Files (*.xlsx;*.xlsm; *.csv)"
    );

    if(!fileName_NewReq.isEmpty() && !fileName_NewReq.isNull())
    {
        ui->lineEdit_NewReq->setText(fileName_NewReq);
        newReqFileLastPath = QFileInfo(fileName_NewReq).path();
        if (QFileInfo(fileName_NewReq).suffix().toLower().contains("xls"))
        {
            QXlsx::Document  docToListSheets(fileName_NewReq);
            foreach(QString name, docToListSheets.sheetNames())
            {
                //qDebug() << "new sheets" <<  name;
                ui->comboBox_SheetsNew->addItem(name);
            }

        }

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
    ui->comboBox_SheetsNew->clear();
    asNewSheetName.clear();


}

void MainWindow::on_btnOldReq_clicked()
{
    EmptyChangesTable();
    ui->comboBox_SheetsOld->clear();
    //delete sheet names
    asOldSheetName.clear();

    // local only, it is use later from edit window


    QString fileName_OldReq = QFileDialog::getOpenFileName
    (
      this,
      "Open Old Requirements",
      oldReqFileLastPath,
      "Excel Files (*.xlsx;*.xlsm; *.csv)"
    );

    if(!fileName_OldReq.isEmpty() && !fileName_OldReq.isNull())
    {
        ui->lineEdit_OldReq->setText(fileName_OldReq);
        oldReqFileLastPath = QFileInfo(fileName_OldReq).path();
        //load sheets
        if (QFileInfo(fileName_OldReq).suffix().toLower().contains("xls"))
        {
           QXlsx::Document  docToListSheets(fileName_OldReq);
           foreach(QString name, docToListSheets.sheetNames())
            {
               //qDebug() << "old sheets" <<  name;
               ui->comboBox_SheetsOld->addItem(name);
            }

        }
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
      ui->comboBox_SheetsOld->clear();
      asOldSheetName.clear();

}

void MainWindow::on_btnCompare_clicked()
{
    //set visible "compare" page in result tab
    ui->tabWidget_Results->setCurrentWidget(ui->pageQuery);

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
    //qDebug() << __LINE__;
    QXlsx::Document  newReqDoc(fileName_NewReq);
    //qDebug() << __LINE__;

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


    if (ui->comboBox_SheetsNew->count() > 1)   //if only one, then it is default...
    {
        newReqDoc.selectSheet(ui->comboBox_SheetsNew->currentText());
    }
    asNewSheetName = newReqDoc.currentSheet()->sheetName();


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



    if (ui->comboBox_SheetsOld->count() > 1)   //if only one, then it is default...
    {
        oldReqDoc.selectSheet(ui->comboBox_SheetsOld->currentText());
    }
    asOldSheetName = oldReqDoc.currentSheet()->sheetName();


    //save paths to ini file
    iniSettings->setValue("paths/newReqFileLastPath", newReqFileLastPath);
    iniSettings->setValue("paths/oldReqFileLastPath", oldReqFileLastPath);


    //prepare report file
    fileName_Report =   QFileInfo(fileName_NewReq).path() + "/" +
                        QFileInfo(fileName_NewReq).completeBaseName() +
                        "_to_" +
                        QFileInfo(fileName_OldReq).completeBaseName();




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

    //close detail
    manageStoredQueries->close();


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



    QStringList lstNewHeaders;
    lstNewHeaders.clear();
    QMap<QString, int> mapNewHeaders;
    mapNewHeaders.clear();

    for(int col= kn_ReqIDCol; col <= newLastCol; col++)
    {
      QString asTemp = getCellValue(kn_HeaderRow, col, newReqDoc, true);
      asTemp.replace("."," ");
      lstNewHeaders << asTemp;
      mapNewHeaders[asTemp] = col;
    }

        //Test Columns
        foreach (QString asTemp, lstNewHeaders)
        {
          //qDebug() << "new: " << asTemp;
        }


    QStringList lstOldHeaders;
    lstOldHeaders.clear();
    QMap<QString, int> mapOldHeaders;
    mapOldHeaders.clear();

    for(int col= kn_ReqIDCol; col <= oldLastCol; col++)
    {
        QString asTemp = getCellValue(kn_HeaderRow, col, oldReqDoc, true);
        asTemp.replace("."," ");
        lstOldHeaders << asTemp;
        mapOldHeaders[asTemp] = col;
    }

    //Test Columns
    foreach (QString asTemp, lstOldHeaders)
    {
      //qDebug() << "old: " << asTemp;
    }







    QStringList lstNewReqIDs;
    lstNewReqIDs.clear();
    QStringList lstOldReqIDs;
    lstOldReqIDs.clear();


    db_Req.close();
    db_Req.open();

    //Create database table OLD
    QSqlQuery dbQuery(db_Req);
    modelOld->revertAll();
    modelOld->clear();
    modelNew->revertAll();
    modelNew->clear();
    modelUserQuery->clear();


    ui->tableView_Old->setSortingEnabled(false);
    ui->listWidget_oldColumns->clear();
    bool boInsertedOK = true;

    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tOLD`");
    if(!boGenResult) boInsertedOK = false;
    //qDebug() <<"Old Drop" << boGenResult << dbQuery.lastError().text();


    boGenResult = dbQuery.exec("CREATE TABLE `tOLD` (rowid INTEGER PRIMARY KEY)");
    if(!boGenResult) boInsertedOK = false;
    //qDebug() <<"Old Create" << boGenResult << dbQuery.lastError().text();

    int iOldColDeb = 0;
    foreach (QString asColumnName, lstOldHeaders)
    {
       boGenResult = dbQuery.exec("ALTER TABLE `tOLD` ADD \""+asColumnName+"\" VARCHAR" );
       if(!boGenResult) boInsertedOK = false;
       //qDebug() <<"Old Add Col" << asColumnName << boGenResult << dbQuery.lastError().text();
       ui->listWidget_oldColumns->addItem(asColumnName);
       iOldColDeb++;
    }


    for (int oldRow = kn_FistDataRow; oldRow <= oldLastRow; ++oldRow)
    {
        QString asColumns = "";
        QString asValues  = "";
        foreach (QString asColumnName, lstOldHeaders)
        {
          asColumns += "`"+asColumnName+"`";
          asColumns +=", ";
          QString asValue = getCellValueValue(oldRow, mapOldHeaders[asColumnName], oldReqDoc, true);
          if(ui->cb_IngnoreWaC->isChecked()) asValue = ignoreWhiteAndCase(asValue);
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
    modelOld->setTable("tOLD");
    modelOld->setFilter("");
    modelOld->select();
    //qDebug() << "modelOld->rowCount()" << modelOld->rowCount() << iOldColDeb << oldLastRow;

    ui->tableView_Old->setModel(modelOld);

//    //adapt column with OLD
//    ui->tableView_Old->resizeColumnsToContents();
//    for(int i=0; i<ui->tableView_Old->model()->columnCount(); i++)
//    {
//     if(ui->tableView_Old->columnWidth(i) > knMAX_VIEW_COLUMN_WIDTH)
//       ui->tableView_Old->setColumnWidth(i, knMAX_VIEW_COLUMN_WIDTH);
//    }
     ui->tableView_Old->setSortingEnabled(true);



    //Create database table NEW
    ui->tableView_New->setSortingEnabled(false);
    ui->listWidget_newColumns->clear();
    boInsertedOK = true;


    boGenResult = dbQuery.exec("DROP TABLE IF EXISTS `tNEW`");
    if(!boGenResult) boInsertedOK = false;
    //qDebug() << "New Drop" << boGenResult << dbQuery.lastError().text();



    boGenResult = dbQuery.exec("CREATE TABLE `tNEW` (rowid INTEGER PRIMARY KEY)");
    if(!boGenResult) boInsertedOK = false;
    //qDebug() <<"New Create" << boGenResult << dbQuery.lastError().text();


    int iNewColDeb = 0;
    foreach (QString asColumnName, lstNewHeaders)
    {
       boGenResult = dbQuery.exec("ALTER TABLE `tNEW` ADD \""+asColumnName+"\" VARCHAR" );
       if(!boGenResult) boInsertedOK = false;
       //qDebug() <<"New Add Col" << asColumnName << boGenResult << dbQuery.lastError().text();
       ui->listWidget_newColumns->addItem(asColumnName);
       iNewColDeb++;
    }


    //insert data NEW
    for (int newRow = kn_FistDataRow; newRow <= newLastRow; ++newRow)
    {
      QString asColumns = "";
      QString asValues  = "";
      foreach (QString asColumnName, lstNewHeaders)
      {
        asColumns += "`"+asColumnName+"`";
        asColumns +=", ";
        QString asValue = getCellValueValue(newRow, mapNewHeaders[asColumnName], newReqDoc, true);
        if(ui->cb_IngnoreWaC->isChecked()) asValue = ignoreWhiteAndCase(asValue);
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
    modelNew->setTable("tNEW");
    modelNew->setFilter("");
    modelNew->select();
    //qDebug() << "modelNew->rowCount()" << modelNew->rowCount() << iNewColDeb << newLastRow;
    ui->tableView_New->setModel(modelNew);

//    ui->tableView_New->resizeColumnsToContents();
//    for(int i=0; i<ui->tableView_New->model()->columnCount(); i++)
//    {
//      if(ui->tableView_New->columnWidth(i) > knMAX_VIEW_COLUMN_WIDTH)
//        ui->tableView_New->setColumnWidth(i, knMAX_VIEW_COLUMN_WIDTH);
//    }
    ui->tableView_New->setSortingEnabled(true);


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



void MainWindow::on_btn_Debug_clicked()
{
    //if(detailView != nullptr) detailView->exec();
    //detailView->show();
    //detailView->setTexts("My Old", "My New");
    manageStoredQueries->show();
    qDebug() << ui->comboBox_SheetsOld->count()
             << ui->comboBox_SheetsOld->currentIndex()
             << ui->comboBox_SheetsOld->currentText();
    qDebug() << ui->comboBox_SheetsNew->count()
             << ui->comboBox_SheetsNew->currentIndex()
             << ui->comboBox_SheetsNew->currentText();
}












void MainWindow::on_btn_ExecQuery_clicked()
{

    asQueryTextToReport.clear();
    ui->tableView_Query->setSortingEnabled(false);

    QString asQueryText = ui->plainTextEdit_Query->toPlainText();
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
      asQueryTextToReport = asQueryText;
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
    //enable export
    ui->btnWrite->setEnabled(true);
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
  ui->plainTextEdit_Query->insertPlainText(asToAdd + " ");
}

void MainWindow::on_btn_ClearQueryLine_clicked()
{
   ui->plainTextEdit_Query->clear();
}

void MainWindow::on_btn_ManageQueries_clicked()
{
    manageStoredQueries->asCurrentQuery = ui->plainTextEdit_Query->toPlainText();
    manageStoredQueries->exec();
    if(!manageStoredQueries->asQueryToLoad.isEmpty())
    {
       ui->plainTextEdit_Query->setPlainText(manageStoredQueries->asQueryToLoad);
    }
}
