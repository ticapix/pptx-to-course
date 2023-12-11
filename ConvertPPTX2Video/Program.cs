using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using Microsoft.Office.Core;
// using Microsoft.Office.Interop.Graph;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using FFMpegCore;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace PPTX2Course
{
    class App
    {
        readonly static string rootFolder = "C:\\Users\\ticap\\Documents\\PPTX to courses\\tests\\";
        static void Main(string[] args)
        {
            Console.WriteLine("The current time is " + DateTime.Now);
            TestFile("slides no transition no animation.pptx");
            TestFile("slide animation 01.pptx");
            TestFile("slide animation 02.pptx");
        }

        static private void TestFile(string pptxFileName) {
            string videoFileName = App.rootFolder + "output.mp4";
            pptxFileName = rootFolder + pptxFileName;
            int defaultTransitionDurationMs = 4000;
            TimeSpan pptxTimeSpan = PPTXInfo.GetSlideDurations(pptxFileName, defaultTransitionDurationMs);
            PPTX2Video.ConvertToVideo(pptxFileName, videoFileName, defaultTransitionDurationMs);
            var mediaInfo = FFProbe.Analyse(videoFileName);
            Console.WriteLine($"Created video duration: {mediaInfo.Duration}.");
            int maxDeltaMs = 150;
            Debug.Assert(Math.Abs((pptxTimeSpan - mediaInfo.Duration).Milliseconds) < maxDeltaMs, $"{pptxFileName}: the difference between the computed duration {pptxTimeSpan} and rendered duration {mediaInfo.Duration} is greater than {maxDeltaMs} ms.");
        }
    }

    class PPTX2Video {
// https://headontech.wordpress.com/2017/01/10/convert-microsoft-powerpoint-presentation-to-a-video-using-c-net-and-hosting-tips-iis/
        public static void ConvertToVideo(string pptxFileName, string mp4FileName, int defaultTransitionDurationMs = 5000) {
            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = MsoTriState.msoTrue;
            ppApp.WindowState = PpWindowState.ppWindowMinimized;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(pptxFileName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            System.Threading.Thread.Sleep(180);
            oPres.UpdateLinks();
            try {
                oPres.UpdateLinks();
                //CreateVideo(string FileName, bool UseTimingsAndNarrations, int DefaultSlideDuration, int VertResolution, int FramesPerSecond, int Quality)
                oPres.CreateVideo(mp4FileName, true, defaultTransitionDurationMs/1000, 480, 30, 85);

                while (oPres.CreateVideoStatus == Microsoft.Office.Interop.PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusInProgress || oPres.CreateVideoStatus == Microsoft.Office.Interop.PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusQueued) {
                    System.Threading.Thread.Sleep(1000);
                }
                Console.WriteLine("Video is Created !!");
            }
            catch (Exception er) {
                Console.WriteLine($"ERROR: {er.StackTrace}");
            }
            finally {
                oPres.Close();
                ppApp.Quit();
                GC.Collect();
            }
        }
    }

    class PPTXInfo
    {
        public static TimeSpan GetSlideDurations(string powerPointFileName, int defaultTransitionDuration = 5000)
        {
            try
            {
                TimeSpan presentationDuration = TimeSpan.FromSeconds(0);
                using (PresentationDocument pptDocument = PresentationDocument.Open(powerPointFileName, false))
                {
                    PresentationPart presentationPart = pptDocument.PresentationPart;
                    DocumentFormat.OpenXml.Presentation.Presentation presentation = presentationPart.Presentation;
                    int slideNumber = 1;
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        DocumentFormat.OpenXml.Presentation.Slide slide = slidePart.Slide;
                        var advanceAfterTimeDuration = PPTXComputeDelay.GetSlideAdvanceAfterTimeDuration(slide, defaultTransitionDuration); // duration for the slide to be displayed
                        var animationsDuration = PPTXComputeDelay.GetSlideAnimationsDuration(slide); // duration of the animation + ??
                        var transitionDuration = PPTXComputeDelay.GetSlideTransitionsDuration(slide); // duration of the transition effect between slides
                        var totalSlideDuration = Math.Max(advanceAfterTimeDuration, animationsDuration) + transitionDuration;
                        TimeSpan slideTime = TimeSpan.FromMilliseconds(totalSlideDuration);
                        presentationDuration = presentationDuration.Add(slideTime);

                        Console.WriteLine($"Slide {slideNumber} Total Duration: {totalSlideDuration} ms. (aat: {advanceAfterTimeDuration} ms, ani: {animationsDuration} ms trn: {transitionDuration} ms)");
                        slideNumber++;
                    }
                    Console.WriteLine($"Total Presentation Duration: {presentationDuration}.");
                    return presentationDuration;

                    // Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Problem occurred parsing file {powerPointFileName}.  Exception: {ex}");
            }
        }
    }
}
