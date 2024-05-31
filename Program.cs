using System;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.WebUtilities;
using Newtonsoft.Json;
using ScipBe.Common.Office.OneNote;
using System.Timers;
using System.Threading;
using System.CommandLine;
using System.CommandLine.Invocation;
using Microsoft.Office.Interop.OneNote;


class Program
{
    static System.Timers.Timer timer;
    static RootCommand rootCommand;

    static async Task Main(string[] args)
    {
        rootCommand = new RootCommand();
        var waitingTimeOption = new Option<int>(
            new string[] { "--time", "-t" },
            () => 0,
            "待命时间（秒），默认为0"
        );
        var portOption = new Option<int>(
            new string[] { "--port", "-p" },
            () => 8022,
            "监听端口，默认 8022"
       );

        // 创建根命令并添加选项
        rootCommand.AddOption(waitingTimeOption);
        rootCommand.AddOption(portOption);
        Func<int, int, Task> action = async (time, port) =>
                {
                    using (Mutex mutex = new Mutex(true, "OneNoteSearcher", out bool createdNew))
                    {
                        if (createdNew)
                        {
                            // 设置 HTTP 服务监听地址和端口号
                            string ipAddress = "127.0.0.1";

                            // 创建 HttpListener 以监听 HTTP 连接
                            HttpListener listener = new HttpListener();
                            listener.Prefixes.Add($"http://{ipAddress}:{port}/");
                            listener.Start();
                            Console.WriteLine($"HTTP server started on {ipAddress}:{port}");

                            if (time > 0)
                            {
                                time *= 1000;
                                timer = new System.Timers.Timer(time);
                                timer.Elapsed += TimerElapsed;
                                timer.AutoReset = false;
                                timer.Start();
                            }

                            while (true)
                            {
                                // 接受客户端连接并处理请求
                                var context = await listener.GetContextAsync();
                                _ = Task.Run(() => ProcessHttpRequest(context, time)); // 使用 Task.Run 启动新的线程处理请求
                            }
                        }
                        else
                        {
                            // 如果互斥体已存在，则关闭当前实例
                            Console.WriteLine("Another instance is already running. Exiting...");
                            Thread.Sleep(2000); // 为了让用户看到提示信息，程序暂停2秒钟
                        }
                    }
                };
        rootCommand.Handler = CommandHandler.Create<int, int>(action);
        await rootCommand.InvokeAsync(args);
    }

    private static void TimerElapsed(object sender, ElapsedEventArgs e)
    {
        Thread.Sleep(2000);
        Console.WriteLine("Timeout, The instance will exit after 2 seconds");
        Environment.Exit(0);
    }

    static async Task ProcessHttpRequest(HttpListenerContext context, int waiting)
    {
        if (waiting > 0)
        {
            timer.Stop();
            timer.Interval = waiting;
            timer.Start();
        }

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
        else if (path.StartsWith("/open"))
        {
            if (context.Request.HttpMethod == "GET")
            {
                responseMessage = HandleOpenRequest(context);
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

    static string HandleOpenRequest(HttpListenerContext context)
    {
        // 获取 URL 参数
        string query = context.Request.Url.Query;
        var queryParameters = QueryHelpers.ParseQuery(query);
        string id = queryParameters.ContainsKey("id") ? queryParameters["id"].ToString() : null;

        if (string.IsNullOrEmpty(id))
        {
            context.Response.StatusCode = (int)HttpStatusCode.BadRequest;
            return "Missing id parameter";
        }

        Console.WriteLine($"Open {id}");
        Microsoft.Office.Interop.OneNote.Application oneNote;
        oneNote = new Microsoft.Office.Interop.OneNote.Application();
        oneNote.NavigateTo(id);
        return "Succeed";
    }
}
