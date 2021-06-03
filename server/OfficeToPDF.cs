using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using pdftron;
using pdftron.Common;
using pdftron.PDF;
using pdftron.SDF;

// This example shows how to create Azure functions using PDFNet SDK.
// A REST API request was posted with base64 encoded data by the client.
// The request would be processed by the server and a response with base64 encoded data of OfficeToPDF output would be sent to the client.
namespace FunctionApp
{
    public static class OfficeToPDF
    {

        [FunctionName("OfficeToPDF")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            try
            {
                log.LogInformation("C# HTTP trigger function processed a request.");
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                dynamic data = JsonConvert.DeserializeObject(requestBody);

                PDFNet.Initialize();

                string file_data = "";
                string json_data = "";
                log.LogInformation("Reading file data from the request ...");
                try
                {
                    file_data = System.Convert.ToString(data.file.data);
                    if (!string.IsNullOrEmpty(file_data))
                    {
                        log.LogInformation($"data: {data}");
                        log.LogInformation("Converting base64 string to bytes...");
                        Byte[] input_bytes = System.Convert.FromBase64String(file_data);
                        Byte[] output_bytes;
                        log.LogInformation("Processing using OfficeToPDF() ...");
                        pdftron.Filters.MemoryFilter memoryFilter = new pdftron.Filters.MemoryFilter(input_bytes.Length, false);
                        pdftron.Filters.FilterWriter writer = new pdftron.Filters.FilterWriter(memoryFilter);
                        writer.WriteBuffer(input_bytes);
                        writer.Flush();
                        memoryFilter.SetAsInputFilter();
                        PDFDoc pdfdoc = new PDFDoc();
                        pdftron.PDF.Convert.OfficeToPDF(pdfdoc, memoryFilter, null);

                        log.LogInformation("Saving output as bytes ...");
                        output_bytes = pdfdoc.Save(SDFDoc.SaveOptions.e_linearized);

                        log.LogInformation("Converting output bytes to base64 string and send a response ...");
                        string base64_str = System.Convert.ToBase64String(output_bytes);
                        // create json data
                        var myData = new
                        {
                            type = "File",
                            title = "Transfer base64 encoded Office2PDF output.",
                            file = new
                            {
                                encoding = "base64",
                                data = base64_str,
                                fileName = "docx2pdf.pdf",
                                contentType = "application/pdf"
                            }
                        };
                        // transform it to Json object
                        json_data = JsonConvert.SerializeObject(myData);
                        log.LogInformation($"data: {myData}");
                    }
                }
                catch (Exception ex)
                {
                    log.LogInformation("No data is sent!");
                }
                log.LogInformation("Sending response to the request ...");
                string responseMessage = !string.IsNullOrEmpty(json_data) ? json_data : $"Hello! PDFNet version = {PDFNet.GetVersion()}. This HTTP triggered function executed successfully.";
                return new OkObjectResult(responseMessage);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error occured: {ex.Message}");
                return new OkObjectResult("Exception occurred!");
            }
           
        }
    }
}

