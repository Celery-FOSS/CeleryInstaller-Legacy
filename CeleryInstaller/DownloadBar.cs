using System.Net;
using System.Security.Authentication;
using Spectre.Console;

namespace CeleryInstaller.Utils;

public sealed class DownloadBar {
    public HttpClient _internalHttpClient;
    private const int chunkSize = 32768;

    public DownloadBar() {
        _internalHttpClient = new HttpClient(new HttpClientHandler() {
            SslProtocols = SslProtocols.Tls12,
            UseCookies = false,
            AutomaticDecompression = DecompressionMethods.All
        });
    }

    public DownloadBar(HttpClient httpClient) {
        _internalHttpClient = httpClient;
    }

    public async Task<IDictionary<Stream, DownloadBarItem>> StartDownloadBar(IEnumerable<DownloadBarItem> urlList,
        CancellationToken token = default) {
        urlList = urlList.ToList(); // Enumerate all objects, avoids possible multiple enumeration.
        var urlCount = urlList.Count();
        Dictionary<HttpResponseMessage, DownloadBarItem> streamTracker = new(urlCount);

        for (var i = 0; urlList.Skip(i).Any(); i++) {
            var iter = i;
            streamTracker.Add(
                await _internalHttpClient
                     .GetAsync(urlList.ElementAt(iter).Url, HttpCompletionOption.ResponseHeadersRead, token)
                     .ConfigureAwait(false),
                urlList.ElementAt(i)
            );
        }

        Dictionary<Stream, DownloadBarItem> FinalizedDownloads = new(urlCount);

        await AnsiConsole.Progress()
                         .AutoRefresh(true) // Turn off auto refresh
                         .AutoClear(false) // Do not remove the task list when done
                         .HideCompleted(false) // Hide tasks as they are completed
                         .StartAsync(async ctx => {
                              var progTasks = new ProgressTask[streamTracker.Count + 1];
                              var threadList = new Thread[streamTracker.Count + 1];

                              for (var i = 0; i < streamTracker.Count; i++) {
                                  var currIterator = i;
                                  var dlName = urlList.ElementAt(i).ItemName.RemoveMarkup();
                                  var contentLength =
                                      (double)streamTracker.ElementAt(i).Key.Content.Headers.ContentLength!;
                                  progTasks[i] =
                                      ctx.AddTask(
                                          $"[green]Downloading [yellow bold]{dlName}[/]...[/]",
                                          true,
                                          contentLength
                                      );

                                  Thread thread = new(() => {
                                      var x = streamTracker.ElementAt(currIterator);
                                      var y = urlList.ElementAt(currIterator);
                                      ProgressUpdateCode(currIterator, contentLength, progTasks[currIterator], FinalizedDownloads, x, y, token).GetAwaiter().GetResult();
                                  });
                                  thread.Start();
                                  threadList[i] = thread;
                              }

                              while (!ctx.IsFinished)
                                  await Task.Delay(100, token).ConfigureAwait(false);
                          }).ConfigureAwait(false);

        return FinalizedDownloads;
    }

    // I did this to avoid a possible program-ending exception.
    private static async Task ProgressUpdateCode(
        int index,
        double lengthOfContent,
        ProgressTask trackedTask,
        Dictionary<Stream, DownloadBarItem> FinalizedDownloads,
        KeyValuePair<HttpResponseMessage, DownloadBarItem> streamTracker,
        DownloadBarItem item,
        CancellationToken token) {
        var httpStream =
            await streamTracker.Key.Content.ReadAsStreamAsync(token).ConfigureAwait(false);
        var finished = File.Create(streamTracker.Value.SavePath);
        int readData = 0, lastRead;
        var buffer = new byte[chunkSize];

        do {
            lastRead = await httpStream.ReadAsync(buffer.AsMemory(0, chunkSize), token).ConfigureAwait(false);

            await finished.WriteAsync(buffer.AsMemory(0, lastRead), token).ConfigureAwait(false);

            readData += lastRead;

            trackedTask.Value = readData;
        } while (readData < lengthOfContent);

        await finished.FlushAsync(token).ConfigureAwait(false); // Save all data to file

        await finished.DisposeAsync();

        FinalizedDownloads.Add(finished, item);
    }
}