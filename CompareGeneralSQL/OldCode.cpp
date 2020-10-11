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


ui->cb_IngnoreWaC->setChecked(false);




