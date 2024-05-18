using System;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.WebUtilities;
using Newtonsoft.Json;
using ScipBe.Common.Office.OneNote;

class Program
{
    static async Task Main(string[] args)
    {
        // 设置 HTTP 服务监听地址和端口号
        string ipAddress = "127.0.0.1";
        int port = 12345;

        // 创建 HttpListener 以监听 HTTP 连接
        HttpListener listener = new HttpListener();
        listener.Prefixes.Add($"http://{ipAddress}:{port}/");
        listener.Start();
        Console.WriteLine($"HTTP server started on {ipAddress}:{port}");

        while (true)
        {
            // 接受客户端连接并处理请求
            var context = await listener.GetContextAsync();
            Task.Run(() => ProcessHttpRequest(context)); // 使用 Task.Run 启动新的线程处理请求
        }
    }

    static async Task ProcessHttpRequest(HttpListenerContext context)
    {
        // 获取请求路径
        string path = context.Request.Url.AbsolutePath;
        string responseMessage = string.Empty;

        if (path.StartsWith("/search"))
        {
            if (context.Request.HttpMethod == "GET")
            {
                responseMessage = HandleSearchRequest(context);
            }
            else
            {
                context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
                responseMessage = "Method not allowed";
            }
        }
        else if (path == "/open")
        {
            if (context.Request.HttpMethod == "POST")
            {
                responseMessage = await HandleOpenRequest(context);
            }
            else
            {
                context.Response.StatusCode = (int)HttpStatusCode.MethodNotAllowed;
                responseMessage = "Method not allowed";
            }
        }
        else
        {
            context.Response.StatusCode = (int)HttpStatusCode.NotFound;
            responseMessage = "Invalid API endpoint";
        }

        // 设置响应状态码和内容类型
        context.Response.StatusCode = (int)HttpStatusCode.OK;
        context.Response.ContentType = "application/json";

        // 发送响应
        byte[] buffer = Encoding.UTF8.GetBytes(responseMessage);
        context.Response.ContentLength64 = buffer.Length;
        await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
        context.Response.OutputStream.Close();

        Console.WriteLine("Response sent.");
    }

    static string HandleSearchRequest(HttpListenerContext context)
    {
        // 获取 URL 参数
        string query = context.Request.Url.Query;
        var queryParameters = QueryHelpers.ParseQuery(query);
        string keyword = queryParameters.ContainsKey("keyword") ? queryParameters["keyword"].ToString() : null;

        if (string.IsNullOrEmpty(keyword))
        {
            context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
            return "Missing keyword parameter";
        }

        Console.WriteLine($"Searching for {keyword}");
        var pages = OneNoteProvider.FindPages(keyword);
        return JsonConvert.SerializeObject(pages);
    }

    static async Task<string> HandleOpenRequest(HttpListenerContext context)
    {
        string requestData;
        using (var reader = new System.IO.StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
        {
            requestData = await reader.ReadToEndAsync();
        }

        Console.WriteLine("Performing open operation with data: " + requestData);

        // 在这里添加具体的逻辑来处理打开操作
        // 示例操作：
        // string result = OpenOperation(requestData);
       

        return "Open action performed";
    }
}
