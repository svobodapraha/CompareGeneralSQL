


1. QSqlQuery query("SELECT * FROM people");
2. int idName = query.record().indexOf("name");
3. while (query.next())
{
4. QString name = query.value(idName).toString();
5. qDebug() << name;
}


http://katecpp.github.io/sqlite-with-qt/


int sheetIndexNumber = 0;
foreach( QString currentSheetName, xlsxDoc.sheetNames() )
{
	// get current sheet 
	AbstractSheet* currentSheet = xlsxDoc.sheet( currentSheetName );
	if ( NULL == currentSheet )
		continue;

	// get full cells of current sheet
	int maxRow = -1;
	int maxCol = -1;
	currentSheet->workbook()->setActiveSheet( sheetIndexNumber );
	Worksheet* wsheet = (Worksheet*) currentSheet->workbook()->activeSheet();
	if ( NULL == wsheet )
		continue;

	QString strSheetName = wsheet->sheetName(); // sheet name
	qDebug() << strSheetName; 

	QVector<CellLocation> clList = wsheet->getFullCells( &maxRow, &maxCol );

	QVector< QVector<QString> > cellValues;
	for (int rc = 0; rc < maxRow; rc++)
	{
		QVector<QString> tempValue;
		for (int cc = 0; cc < maxCol; cc++)
		{
			tempValue.push_back(QString(""));
		}
		cellValues.push_back(tempValue);
	}

	for ( int ic = 0; ic < clList.size(); ++ic )
	{
		CellLocation cl = clList.at(ic); // cell location

		int row = cl.row - 1;
		int col = cl.col - 1;

		QSharedPointer<Cell> ptrCell = cl.cell; // cell pointer

		// value of cell
		QVariant var = cl.cell.data()->value();
		QString str = var.toString();

		cellValues[row][col] = str;
	}

	for (int rc = 0; rc < maxRow; rc++)
	{
		for (int cc = 0; cc < maxCol; cc++)
		{
			QString strCell = cellValues[rc][cc];
			qDebug() << "( row : " << rc 
			         << ", col : " << cc 
			         << ") " << strCell; // display cell value
		}
	}

	sheetIndexNumber++;
}  
