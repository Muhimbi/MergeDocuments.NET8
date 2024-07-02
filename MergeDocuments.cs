using PDFA_Conversion;
using ServiceReference2;
using System.ServiceModel;

namespace MergeDocumentsSample
{
    class MergeDocuments{
        static void Main(string[] args){
            DocumentConverterServiceClient? client = null;
            string[] sourceFiles = { "C:\\Converter\\test1.pdf", "C:\\Converter\\test2.docx" };

            try{
                Console.WriteLine("Merging...");
                ProcessingOptions processingOptions = getProcessingOptions(sourceFiles);
                client = UtilClass.OpenService();
                BatchResults results = client.ProcessBatch(processingOptions);
                byte[] mergedFile = results.Results[0].File;
                /* string hour = DateTime.Now.Hour.ToString();
                 string minute = DateTime.Now.Minute.ToString();
                 string second = DateTime.Now.Second.ToString(); */

                //  string destinationFileName = "C:\\Converter\\MergedOutput"+hour+minute+second+".pdf";
                string destinationFileName = "C:\\Converter\\MergedOutput.pdf";
                using (FileStream fs = File.Create(destinationFileName))
                {
                    fs.Write(mergedFile, 0, mergedFile.Length);
                    fs.Close();
                }
                Console.WriteLine("...to "+ destinationFileName);
            }

            catch (FaultException<WebServiceFaultException> ex){
                Console.WriteLine("FaultException occurred: ExceptionType: " + ex.Detail.ExceptionType.ToString());
            }

            catch (Exception ex){
                Console.WriteLine(ex.ToString());
            }

            finally{
                if (client != null)
                    UtilClass.CloseService(client);
            }
        }

        private static ProcessingOptions getProcessingOptions(string[] sourceFileNames) {
            ProcessingOptions processingOptions = new ProcessingOptions();

            MergeSettings mergeSettings = new MergeSettings();
            mergeSettings.BreakOnError = false;
            processingOptions.MergeSettings = mergeSettings;

            List<SourceFile> sourceFileList = new List<SourceFile>();
            foreach (string sourceFileName in sourceFileNames) {
                byte[] fileContent = File.ReadAllBytes(sourceFileName);

                OpenOptions openOptions = new OpenOptions();
                openOptions.OriginalFileName = Path.GetFileName(sourceFileName);
                openOptions.FileExtension = Path.GetExtension(sourceFileName);

                ConversionSettings conversionSettings = new ConversionSettings();
                conversionSettings.Fidelity = ConversionFidelities.Full;
                conversionSettings.Quality = ConversionQuality.OptimizeForPrint;

                FileMergeSettings fileMergeSettings = new FileMergeSettings();
                fileMergeSettings.TopLevelBookmark = openOptions.OriginalFileName;

                SourceFile sourceFile = new SourceFile();
                sourceFile.OpenOptions = openOptions;
                sourceFile.ConversionSettings = conversionSettings;
                sourceFile.MergeSettings = fileMergeSettings;
                sourceFile.File = fileContent;
                sourceFileList.Add(sourceFile);
                Console.WriteLine(sourceFileName);
            }

            processingOptions.SourceFiles = sourceFileList.ToArray();
            return processingOptions;
        }
    }
}