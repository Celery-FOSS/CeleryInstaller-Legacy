using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using CeleryInstaller.Utils;
using Microsoft.Win32;
using RedistributableChecker;
using Spectre.Console;

namespace CeleryInstaller;

class Program {
    static HttpClient client = new(new HttpClientHandler() {
        SslProtocols = SslProtocols.Tls12,
    });

    private static async Task<string> GetRemoteVersion() {
        return await client.GetStringAsync(
            "https://raw.githubusercontent.com/TheSeaweedMonster/Celery/main/version.txt"); // REPLACE WITH VERSION.TXT URL
    }

    private static async Task<(string, string)> CheckVersion() {
        if (!File.Exists("appversion.txt")) return ("None", await GetRemoteVersion());

        var currentVersion = await File.ReadAllTextAsync("appversion.txt");

        return (currentVersion.Trim(), (await GetRemoteVersion()).Trim());
    }

    static async Task Main(string[] args) {
        DownloadBar dlBar = new();

        var (currentVersion, latestVersion) = await CheckVersion();

        var proceed = false;

        if (currentVersion == "None") {
            AnsiConsole.MarkupLine("No celery found in this folder!");
            AnsiConsole.Markup("Do you wish to install [green]Celery[/]? ");
            proceed = AnsiConsole.Confirm("");
        }
        else {
            AnsiConsole.MarkupLine("Found Celery!");

            if (currentVersion == latestVersion) {
                AnsiConsole.MarkupLine("[green]Celery[/] is [underline]up-to-date[/]!");
            }

            if (currentVersion != latestVersion) {
                AnsiConsole.MarkupLine(
                    $"Latest available version is [green]{latestVersion}[/]. Currently downloaded version is [maroon]{currentVersion}[/].");

                AnsiConsole.Markup("Do you wish to update [green]Celery[/]? ");
                proceed = AnsiConsole.Confirm("");
            }
        }

        if (proceed) {
            AnsiConsole.MarkupLine(currentVersion == "None" ? "Downloading Celery..." : "Updating Celery...");
        }
        else {
            AnsiConsole.Markup("Press any [red]key[/] to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        var bootstrapperPath = Path.Combine(Environment.CurrentDirectory, "Bootstrapper.zip");

        DownloadBarItem downloadItem = new() {
            SavePath = bootstrapperPath,
            ItemName = "Celery's Bootstrapper",
            Url = new Uri(
                "https://cdn.sussy.dev/celery/release.zip"),
        };
        await dlBar.StartDownloadBar(new List<DownloadBarItem>() { downloadItem, });

        AnsiConsole.MarkupLine("The [green]bootstrapper[/] has been downloaded! [yellow]Unpacking[/]...");

        ZipFile.ExtractToDirectory(bootstrapperPath, Environment.CurrentDirectory, true);

        File.Delete(bootstrapperPath);

        AnsiConsole.MarkupLine("[green]Bootstrapper[/] unpacked! Verifying if any pre-requisites are missing...");

        // First == VCRedist
        // Second == Webview2
        var arePrerequisitesInstalled = CheckPrerequisites();

        if (arePrerequisitesInstalled.Item1 && arePrerequisitesInstalled.Item2) {
            AnsiConsole.MarkupLine("All pre-requisites are [green]installed[/]!");
            AnsiConsole.Markup("Press any [red]key[/] to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        AnsiConsole.MarkupLine("[green]Installing[/] missing pre-requisites...");

        if (!arePrerequisitesInstalled.Item2)
            await InstallWebview2();
        if (!arePrerequisitesInstalled.Item1)
            await InstallVCredist_x86();
    }

    private static async Task InstallWebview2() {
        AnsiConsole.MarkupLine("[yellow]Installing[/] Webview 2... This [italic]may[/] take a while.");
        var temporalFile = Path.GetTempFileName();

        var fileHandle = File.Open(temporalFile, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);

        await (await client.GetStreamAsync("https://go.microsoft.com/fwlink/p/?LinkId=2124703"))
           .CopyToAsync(fileHandle);

        await fileHandle.DisposeAsync();

        var proc = Process.Start(new ProcessStartInfo() {
            FileName = temporalFile,
            CreateNoWindow = true,
            Arguments = "/silent /install",
        });

        if (proc == null) {
            AnsiConsole.MarkupLine(
                "Whoops! Something went [red]wrong[/]... We have failed to install Webview2! This is a required dependency for Celery! Please download it and install it manually at [bold yellow link]https://go.microsoft.com/fwlink/p/?LinkId=2124703[/]");
            Environment.Exit(1); // Exit with err.
        }

        try {
            await proc.WaitForExitAsync();

            if (proc.ExitCode == 0) {
                // Everything went well
            }
            else {
                AnsiConsole.MarkupLine(
                    "Whoops! Something went [red]wrong[/]... We have failed to install Webview2! This is a required dependency for Celery! Please download it and install it manually at [bold yellow link]https://go.microsoft.com/fwlink/p/?LinkId=2124703[/]");
                Environment.Exit(1); // Exit with err. 
            }
        }
        finally {
            proc.Dispose();

            File.Delete(temporalFile); // Delete the installer.
        }
    }

    private static async Task InstallVCredist_x86() {
        AnsiConsole.MarkupLine("[yellow]Installing[/] VCRedist x86... This [italic]may[/] take a while.");
        var temporalVcredistPath = Path.GetTempFileName();

        var fileHandle = File.Open(temporalVcredistPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);

        // 2017 redist.
        await (await client.GetStreamAsync("https://aka.ms/vs/17/release/vc_redist.x86.exe")).CopyToAsync(fileHandle);

        await fileHandle.DisposeAsync();

        var proc = Process.Start(new ProcessStartInfo() {
            FileName = temporalVcredistPath,
            Arguments = "/install /passive /silent /norestart",
            CreateNoWindow = true,
        });

        if (proc == null) {
            AnsiConsole.MarkupLine(
                "Whoops! Something went [red]wrong[/]... We have failed to install the Visual C/C++ Redistributable package! This is a required dependency for Celery! Please download it and install it manually at [bold yellow link]https://aka.ms/vs/17/release/vc_redist.x86.exe[/]");
            Environment.Exit(1); // Exit with err.
        }

        try {
            await proc.WaitForExitAsync();
            AnsiConsole.MarkupLine("[green]VCRedist[/] installed successfully!");
        }
        finally {
            proc.Dispose();

            File.Delete(temporalVcredistPath); // Delete the installer.
        }
    }

    private static (bool, bool) CheckPrerequisites() {
        // Check 2015-2022 x86 VCRedist installs and webview 2.
        return (RedistributablePackage.IsInstalled(RedistributablePackageVersion.VC2015to2022x86), VerifyWebview2());
    }

    private static bool VerifyWebview2() {
        switch (RuntimeInformation.OSArchitecture) {
            case Architecture.X86: {
                // Check X86 keys for webview 2.
                var localMachineKey =
                    Registry.LocalMachine.OpenSubKey(
                        "SOFTWARE\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");

                var localUserKey =
                    Registry.CurrentUser.OpenSubKey(
                        "Software\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");

                return localMachineKey != null || localUserKey != null;
            }
            case Architecture.X64: {
                // Check X64 keys for webview 2.
                var localMachineKey =
                    Registry.LocalMachine.OpenSubKey(
                        "SOFTWARE\\WOW6432Node\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");

                var localUserKey =
                    Registry.CurrentUser.OpenSubKey(
                        "Software\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");

                return localMachineKey != null || localUserKey != null;
            }
            default:
                return false;
        }
    }
}