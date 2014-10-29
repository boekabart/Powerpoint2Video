using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Powerpoint2Video
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!args.Any())
                return;
            var fn = Path.GetFullPath(args.First());

            var resolutions = args.Length == 1 ? new[] {"1080"} : args.Skip(1);
            foreach (var res in resolutions)
                CreateVideo(fn, int.Parse(res));
        }

        static bool CreateVideo(string fn, int vertRes)
        {
            var app = new PowerPoint.Application();
            var x = app.Presentations.Open(fn, MsoTriState.msoTrue, WithWindow:MsoTriState.msoFalse);
            var vidFn = Path.ChangeExtension(fn, string.Format(".{0}p.mp4",vertRes));
            Console.WriteLine("Creating {0}", vidFn);
            x.CreateVideo(vidFn, true, 5, vertRes, 30, 100);
            Console.Write("Saving ");
            while (CreatingVideo(x))
            {
                Console.Write('.');
                Thread.Sleep(1000);
            }
            var ok = x.CreateVideoStatus == PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusDone;
            Console.WriteLine(String.Empty);
            Console.WriteLine(ok ? "Done" : "Failed");
            x.Close();
            app.Quit();
            return ok;
        }

        private static bool CreatingVideo(PowerPoint.Presentation presentation)
        {
            switch (presentation.CreateVideoStatus)
            {
                    case PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusQueued:
                    case PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusInProgress:
                    return true;
                    case PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusDone:
                    case PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusFailed:
                    case PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusNone:
                default:
                    return false;
            }
        }
    }
}
