using WordToHtmlConsole;

//After First Run of Code need to set false to oneTimeAPpendHeader = false
bool oneTimeAppendHeader = true;

var htmlFile = "C:\\Users\\user\\Downloads\\WordToHtmlConsole 4\\WordToHtmlConsole\\WordToHtmlConsole\\HtmlFiles\\website.html";
//Append Header from source to destination file
var sourceFilePath = "C:\\Users\\user\\Downloads\\WordToHtmlConsole 4\\WordToHtmlConsole\\WordToHtmlConsole\\SourceFiles\\Employee handbook (1).docx";
//var destinationFilePath = "C:\\Users\\visha\\Downloads\\WordToHtmlConsole\\WordToHtmlConsole\\WordToHtmlConsole\\DestinationFiles\\Document_EmailModule.docx";

var isAddPageBreaker = true;

ProcessWordHeader wordHeader = new ProcessWordHeader();
var isHeaderPresent = wordHeader.DocumentContainsHeader(sourceFilePath);
var isFooterExist = wordHeader.DocumentContainsFooter(sourceFilePath);

if (isAddPageBreaker)
{ 
    wordHeader.AddRandomPageBreak(sourceFilePath);
}

ProcessingModel headerProcessDetails = new ProcessingModel();

  ProcessingModel footerProcessDetails = new ProcessingModel();

if (oneTimeAppendHeader && isHeaderPresent)
{
   headerProcessDetails = wordHeader.AppendHeaderFromSourseToDestinationFile(sourceFilePath);
}

if (isFooterExist)
{
    footerProcessDetails = wordHeader.AppendFooterFromSourseToDestinationFile(sourceFilePath);
}


int noOfHeaderparas = headerProcessDetails.NoOfImages + headerProcessDetails.NoOfHeaderElement;
int noOfFooterparas = footerProcessDetails.NoOfImages + footerProcessDetails.NoOfHeaderElement;
//convert destination file into html to get last paragraph as header of main statement.

ConvertToHtml html = new ConvertToHtml();

var htmlStr = html.ParseDOCX(sourceFilePath,"TestFileName",isHeaderPresent, isFooterExist, noOfHeaderparas,noOfFooterparas);
//var imageStr = html.ParseDOCX(destinationFilePath);
//var images = html.ExtractImages(sourceFilePath);

File.WriteAllText(htmlFile, htmlStr);
//convert destination file into html to get last paragraph as header of main statement.