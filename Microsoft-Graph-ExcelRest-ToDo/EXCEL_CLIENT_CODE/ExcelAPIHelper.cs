//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft_Graph_ExcelRest_ToDo.Models;
using Newtonsoft.Json;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace Microsoft_Graph_ExcelRest_ToDo
{
    public class ExcelApiHelper
    {
        private static string restURLBase = "https://graph.microsoft.com/v1.0/me/drive/items/";
        private static string fileId = null;

        public static async Task LoadWorkbook(string accessToken)
        {
            try
            {
                var fileName = "ToDoList.xlsx";
                var serviceEndpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children";

                String absPath = HttpContext.Current.Server.MapPath("Assets/ToDo.xlsx");
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);


                var filesResponse = await client.GetAsync(serviceEndpoint + "?$select=name,id");

                if (filesResponse.IsSuccessStatusCode)
                {
                    var filesContent = await filesResponse.Content.ReadAsStringAsync();

                    JObject parsedResult = JObject.Parse(filesContent);

                    foreach (JObject file in parsedResult["value"])
                    {

                        var name = (string)file["name"];
                        if (name.Contains("ToDoList.xlsx"))
                        {
                            fileId = (string)file["id"];
                            restURLBase = "https://graph.microsoft.com/v1.0/me/drive/items/" + fileId + "/workbook/worksheets('ToDoList')/";
                            return;
                        }
                    }

                }

                else
                {
                    //Handle failed response
                }

                // We know that the file doesn't exist, so upload it and create the necessary worksheets, tables, and chart.
                var excelFile = File.OpenRead(absPath);
                byte[] contents = new byte[excelFile.Length];
                excelFile.Read(contents, 0, (int)excelFile.Length);
                excelFile.Close();
                var contentStream = new MemoryStream(contents);


                var contentPostBody = new StreamContent(contentStream);
                contentPostBody.Headers.Add("Content-Type", "application/octet-stream");


                // Endpoint for content in an existing file.
                var fileEndpoint = new Uri(serviceEndpoint + "/" + fileName + "/content");

                var requestMessage = new HttpRequestMessage(HttpMethod.Put, fileEndpoint)
                {
                    Content = contentPostBody
                };

                HttpResponseMessage response = await client.SendAsync(requestMessage);

                if (response.IsSuccessStatusCode)
                {
                    //Get the Id of the new file.
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var parsedResponse = JObject.Parse(responseContent);
                    fileId = (string)parsedResponse["id"];
                    restURLBase = "https://graph.microsoft.com/v1.0/me/drive/items/" + fileId + "/workbook/worksheets('ToDoList')/";

                    //Set up workbook and worksheet endpoints
                    var workbookEndpoint = "https://graph.microsoft.com/v1.0/me/drive/items/" + fileId + "/workbook";
                    var worksheetsEndpoint = workbookEndpoint + "/worksheets";

                    //Get session id and add it to the HttpClient's default headers. This will make the changes appear more quickly.
                    var sessionJson = "{" +
                        "'saveChanges': true" +
                        "}";
                    var sessionContentPostbody = new StringContent(sessionJson);
                    sessionContentPostbody.Headers.Clear();
                    sessionContentPostbody.Headers.Add("Content-Type", "application/json");
                    var sessionResponseMessage = await client.PostAsync(workbookEndpoint + "/createsession", sessionContentPostbody);
                    var sessionResponseContent = await sessionResponseMessage.Content.ReadAsStringAsync();
                    JObject sessionObject = JObject.Parse(sessionResponseContent);
                    var sessionId = (string)sessionObject["id"];

                    client.DefaultRequestHeaders.Add("Workbook-Session-Id", sessionId);

                    //Add ToDoList worksheet to the workbook
                    await AddWorksheetToWorkbook("ToDoList", worksheetsEndpoint, client);

                    //Add Summary worksheet to the workbook
                    await AddWorksheetToWorkbook("Summary", worksheetsEndpoint, client);

                    //Add table to ToDoList worksheet
                    await AddTableToWorksheet("ToDoList", "A1:H1", worksheetsEndpoint, client);

                    //Add table too Summary worksheet
                    await AddTableToWorksheet("Summary", "A1:B1", worksheetsEndpoint, client);

                    var patchMethod = new HttpMethod("PATCH");

                    //Rename Table1 in ToDoList worksheet to "ToDoList"
                    var toDoListTableNameJson = "{" +
                            "'name': 'ToDoList'," +
                            "}";

                    var toDoListTableNamePatchBody = new StringContent(toDoListTableNameJson);
                    toDoListTableNamePatchBody.Headers.Clear();
                    toDoListTableNamePatchBody.Headers.Add("Content-Type", "application/json");
                    var toDoListRequestMessage = new HttpRequestMessage(patchMethod, worksheetsEndpoint + "('ToDoList')/tables('Table1')") { Content = toDoListTableNamePatchBody };
                    var toDoListTableNameResponseMessage = await client.SendAsync(toDoListRequestMessage);


                    //Rename ToDoList columns 1-8

                    await RenameColumn("ToDoList", "ToDoList", "Id", "1", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "Title", "2", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "Priority", "3", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "Status", "4", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "PercentComplete", "5", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "StartDate", "6", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "EndDate", "7", worksheetsEndpoint, client);
                    await RenameColumn("ToDoList", "ToDoList", "Notes", "8", worksheetsEndpoint, client);

                    //Rename Summary columns 1-2
                    await RenameColumn("Summary", "2", "Status", "1", worksheetsEndpoint, client);
                    await RenameColumn("Summary", "2", "Count", "2", worksheetsEndpoint, client);


                    //Set numberFormat to text for the two date fields in the ToDoList worksheet table.

                    var dateRangeJSON = "{" +
                        "'numberFormat': '@'" +
                        "}";
                    var datePatchBody = new StringContent(dateRangeJSON);
                    datePatchBody.Headers.Clear();
                    datePatchBody.Headers.Add("Content-Type", "application/json");
                    var dateRequestMessage = new HttpRequestMessage(patchMethod, worksheetsEndpoint + "('ToDoList')/range(address='$F1:$G1000')") { Content = datePatchBody };
                    var dateResponseMessage = await client.SendAsync(dateRequestMessage);


                    //Add three rows to summary table

                    await AddRowToTable("Summary", "2", "Not started", worksheetsEndpoint, client);
                    await AddRowToTable("Summary", "2", "In-progress", worksheetsEndpoint, client);
                    await AddRowToTable("Summary", "2", "Completed", worksheetsEndpoint, client);

                    //Add chart to Summary worksheet
                    var chartJson = "{" +
                        "\"type\": \"Pie\", " +
                        "\"sourcedata\": \"A1:B4\", " +
                        "\"seriesby\": \"Auto\"" +
                        "}";

                    var chartContentPostBody = new StringContent(chartJson);
                    chartContentPostBody.Headers.Clear();
                    chartContentPostBody.Headers.Add("Content-Type", "application/json");
                    var chartCreateResponseMessage = await client.PostAsync(worksheetsEndpoint + "('Summary')/charts/$/add", chartContentPostBody);

                    //Update chart position and title
                    var chartPatchJson = "{" +
                        "'left': 99," +
                        "'name': 'Status'," +
                        "}";

                    var chartPatchBody = new StringContent(chartPatchJson);
                    chartPatchBody.Headers.Clear();
                    chartPatchBody.Headers.Add("Content-Type", "application/json");
                    var chartPatchRequestMessage = new HttpRequestMessage(patchMethod, worksheetsEndpoint + "('Summary')/charts('Chart 1')") { Content = chartPatchBody };
                    var chartPatchResponseMessage = await client.SendAsync(chartPatchRequestMessage);

                    //Close workbook session
                    var closeSessionJson = "{}";
                    var closeSessionBody = new StringContent(closeSessionJson);
                    sessionContentPostbody.Headers.Clear();
                    sessionContentPostbody.Headers.Add("Content-Type", "application/json");
                    var closeSessionResponseMessage = await client.PostAsync(workbookEndpoint + "/closesession", closeSessionBody);

                }

                else
                {
                    //Handle exception

                }

            }

            catch (Exception e)
            {
                //Handle exception

            }
        }

        private static async Task AddWorksheetToWorkbook(string worksheetName, string worksheetsEndpoint, HttpClient client)
        {
            var worksheetJson = "{" +
                            "'name': '" + worksheetName + "'," +
                            "}";

            var worksheetContentPostBody = new StringContent(worksheetJson);
            worksheetContentPostBody.Headers.Clear();
            worksheetContentPostBody.Headers.Add("Content-Type", "application/json");
            var worksheetResponseMessage = await client.PostAsync(worksheetsEndpoint, worksheetContentPostBody);
        }

        private static async Task AddTableToWorksheet(string worksheetName, string tableRange, string worksheetsEndpoint, HttpClient client)
        {
            var tableJson = "{" +
                    "'address': '" + tableRange + "'," +
                    "'hasHeaders': true" +
                    "}";

            var tableContentPostBody = new StringContent(tableJson);
            tableContentPostBody.Headers.Clear();
            tableContentPostBody.Headers.Add("Content-Type", "application/json");
            var tableResponseMessage = await client.PostAsync(worksheetsEndpoint + "('" + worksheetName + "')/tables/$/add", tableContentPostBody);

        }

        private static async Task RenameColumn(string worksheetName, string tableName, string colName, string colNumber, string worksheetsEndpoint, HttpClient client)
        {
            var patchMethod = new HttpMethod("PATCH");
            var colNameJson = "{" +
                    "'values': [['" + colName + "'], [null]] " +
                    "}";

            var colNamePatchBody = new StringContent(colNameJson);
            colNamePatchBody.Headers.Clear();
            colNamePatchBody.Headers.Add("Content-Type", "application/json");
            var colNameRequestMessage = new HttpRequestMessage(patchMethod, worksheetsEndpoint + "('" + worksheetName + "')/tables('" + tableName + "')/Columns('" + colNumber + "')") { Content = colNamePatchBody };
            var colNameResponseMessage = await client.SendAsync(colNameRequestMessage);

        }

        private static async Task AddRowToTable(string worksheetName, string tableName, string rowName, string worksheetsEndpoint, HttpClient client)
        {
            var summaryTableRowJson = "{" +
                    "'values': [['" + rowName + "', '=COUNTIF(ToDoList[PercentComplete],[@Status])']]" +
                "}";
            var summaryTableRowContentPostBody = new StringContent(summaryTableRowJson, System.Text.Encoding.UTF8);
            summaryTableRowContentPostBody.Headers.Clear();
            summaryTableRowContentPostBody.Headers.Add("Content-Type", "application/json");
            var summaryTableRowResponseMessage = await client.PostAsync(worksheetsEndpoint + "('" + worksheetName + "')/tables('" + tableName + "')/rows", summaryTableRowContentPostBody);
        }


        public static async Task<List<ToDoItem>> GetToDoItems(string accessToken)
        {
            List<ToDoItem> todoItems = new List<ToDoItem>();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // New code:
                HttpResponseMessage response = await client.GetAsync(restURLBase + "tables('ToDoList')/Rows");
                if (response.IsSuccessStatusCode)
                {
                    string resultString = await response.Content.ReadAsStringAsync();

                    dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(resultString);
                    JArray y = x.value;

                    todoItems = BuildList(todoItems, y);
                }
            }

            return todoItems;
        }

        private static List<ToDoItem> BuildList(List<ToDoItem> todoItems, JArray y)
        {
            foreach (var item in y.Children())
            {
                var itemProperties = item.Children<JProperty>();

                //Get element that holds row collection
                var element = itemProperties.FirstOrDefault(xx => xx.Name == "values");
                JProperty index = itemProperties.FirstOrDefault(xxx => xxx.Name == "index");

                //The string array of row values
                JToken values = element.Value;

                //LINQ query to get rows from results
                var stringValues = from stringValue in values select stringValue;
                //rows
                foreach (JToken thing in stringValues)
                {
                    IEnumerable<string> rowValues = thing.Values<string>();

                    //Cast row value collection to string array
                    string[] stringArray = rowValues.Cast<string>().ToArray();


                    try
                    {
                        ToDoItem todoItem = new ToDoItem(
                             stringArray[0],
                             stringArray[1],
                             stringArray[3],
                             stringArray[4],
                             stringArray[2],
                             stringArray[5],
                             stringArray[6],
                        stringArray[7]);
                        todoItems.Add(todoItem);
                    }
                    catch (FormatException f)
                    {
                        //Handle exception
                    }
                }
            }

            return todoItems;

        }

        public static async Task<ToDoItem> CreateToDoItem(
                                                 string accessToken,
                                                 string title,
                                                 string priority,
                                                 string status,
                                                 string percentComplete,
                                                 string startDate,
                                                 string endDate,
                                                 string notes)
        {
            ToDoItem newTodoItem = new ToDoItem();

            string id = Guid.NewGuid().ToString();

            var priorityString = "";

            switch (priority)
            {
                case "1":
                    priorityString = "High";
                    break;
                case "2":
                    priorityString = "Normal";
                    break;
                case "3":
                    priorityString = "Low";
                    break;
            }

            var statusString = "";

            switch (status)
            {
                case "1":
                    statusString = "Not started";
                    break;
                case "2":
                    statusString = "In-progress";
                    break;
                case "3":
                    statusString = "Completed";
                    break;
            }
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(restURLBase);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                using (var request = new HttpRequestMessage(HttpMethod.Post, restURLBase))
                {
                    //Create two-dimensional array to hold the row values to be serialized into json
                    object[,] valuesArray = new object[1, 8] { { id, title, percentComplete.ToString(), priorityString, statusString, startDate, endDate, notes } };

                    //Create a container for the request body to be serialized
                    RequestBodyHelper requestBodyHelper = new RequestBodyHelper();
                    requestBodyHelper.index = null;
                    requestBodyHelper.values = valuesArray;

                    //Serialize the final request body
                    string postPayload = JsonConvert.SerializeObject(requestBodyHelper);

                    //Add the json payload to the POST request
                    request.Content = new StringContent(postPayload, System.Text.Encoding.UTF8);


                    using (HttpResponseMessage response = await client.PostAsync("tables('ToDoList')/rows", request.Content))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string resultString = await response.Content.ReadAsStringAsync();
                            dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(resultString);
                        }
                    }
                }
            }
            return newTodoItem;
        }

        public static async Task<FileContentResult> getChartImage(string accessToken)
        {
            FileContentResult returnValue = null;
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/me/drive/items/" + fileId + "/workbook/worksheets('Summary')/");
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                string chartId = null;


                //Take the first chart off the charts collection, because we know there is only one
                HttpResponseMessage chartsResponse = await client.GetAsync("charts");

                var responseContent = await chartsResponse.Content.ReadAsStringAsync();
                var parsedResponse = JObject.Parse(responseContent);
                chartId = (string)parsedResponse["value"][0]["id"];

                HttpResponseMessage response = await client.GetAsync("charts('" + chartId + "')/Image(width=0,height=0,fittingMode='fit')");

                if (response.IsSuccessStatusCode)
                {
                    string resultString = await response.Content.ReadAsStringAsync();

                    dynamic x = JsonConvert.DeserializeObject(resultString);
                    JToken y = x.Last;
                    Bitmap imageBitmap = StringToBitmap(x["value"].ToString());
                    ImageConverter converter = new ImageConverter();
                    byte[] bytes = (byte[])converter.ConvertTo(imageBitmap, typeof(byte[]));
                    returnValue = new FileContentResult(bytes, "image/bmp");
                }
                return returnValue;
            }
        }

        public static Bitmap StringToBitmap(string base64ImageString)
        {
            Bitmap bmpReturn = null;
            byte[] byteBuffer = Convert.FromBase64String(base64ImageString);
            MemoryStream memoryStream = new MemoryStream(byteBuffer);

            memoryStream.Position = 0;

            bmpReturn = (Bitmap)Bitmap.FromStream(memoryStream);
            memoryStream.Close();
            memoryStream = null;
            byteBuffer = null;
            return bmpReturn;


        }

    }
    public class RequestBodyHelper
    {
        public object index;
        public object[,] values;
    }
}