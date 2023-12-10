using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Windows;
using DocumentFormat.OpenXml;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Serialization;

// Hello World! program
namespace PPTX2Course
{
    class App
    {
        static void Main(string[] args)
        {
            //            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            Console.WriteLine("The current time is " + DateTime.Now);
            Console.WriteLine("Hello World!!!");
            TimeSpan timeSpan;
            timeSpan = PPTXInfo.GetSlideDurations("C:\\Users\\ticap\\Downloads\\test animation.pptx");
            Debug.Assert(timeSpan.TotalMilliseconds == 0);
            // PPTXTimingInfo.GetSlideDurations("C:\\Users\\ticap\\Downloads\\test animation.pptx");
        }
    }

    class PPTXInfo
    {
        public static TimeSpan GetSlideDurations(string powerPointFileName)
        {
            try
            {
                TimeSpan presentationDuration = TimeSpan.FromSeconds(0);
                using (PresentationDocument pptDocument = PresentationDocument.Open(powerPointFileName, false))
                {
                    PresentationPart presentationPart = pptDocument.PresentationPart;
                    Presentation presentation = presentationPart.Presentation;
                    int slideNumber = 1;
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        Slide slide = slidePart.Slide;
                        var advanceAfterTimeDuration = PPTXComputeDelay.GetSlideAdvanceAfterTimeDuration(slide); // duration for the slide to be displayed
                        var anitationsDuration = PPTXComputeDelay.GetSlideAnimationsDuration(slide); // duration of the animation + ??
                        var transitionDuration = PPTXComputeDelay.GetSlideTransitionsDuration(slide); // duration of the transition effect between slides

                        var totalSlideDuration = advanceAfterTimeDuration + anitationsDuration + transitionDuration;
                        TimeSpan slideTime = TimeSpan.FromMilliseconds(totalSlideDuration);
                        presentationDuration = presentationDuration.Add(slideTime);

                        Console.WriteLine($"Slide {slideNumber} Total Duration: {totalSlideDuration} ms. (aat: {advanceAfterTimeDuration} ms, ani: {anitationsDuration} ms trn: {transitionDuration} ms)");
                        slideNumber++;
                    }

                    Console.WriteLine($"Total Presentation Duration: {presentationDuration.TotalMilliseconds} msecs.");
                    return presentationDuration;

                    // Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Problem occurred parsing file {powerPointFileName}.  Exception: {ex}");
            }
        }

        // public static string GetSlideTransitionsDuration(SlidePart slidePart)
        // {
        //     string returnDuration = "0";
        //     try
        //     {
        //         Slide slide1 = slidePart.Slide;

        //         var transitions = slide1.Descendants<Transition>();
        //         foreach (var transition in transitions)
        //         {
        //             Console.WriteLine(transition.OuterXml);
        //             if (transition.Duration != null && transition.Duration.HasValue)
        //                 return transition.Duration;
        //             break;
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         //Do nothing
        //     }

        //     return returnDuration;
        // }

        // public static int GetSlideAnimationsDuration(SlidePart slidePart)
        // {
        //     int returnDuration = 0;
        //     try
        //     {
        //         var timing = slidePart.Slide.Timing;
        //         returnDuration = PPTXComputeDelay.ComputeTimingDelay(timing);
        //     }
        //     catch (Exception ex)
        //     {
        //         //Do nothing
        //     }

        //     return returnDuration;
        // }

        // public static string GetSlideAdvanceAfterTimeDuration(SlidePart slidePart)
        // {
        //     string returnDuration = "0";
        //     try
        //     {
        //         Slide slide1 = slidePart.Slide;

        //         var transitions = slide1.Descendants<Transition>();
        //         foreach (var transition in transitions)
        //         {
        //             Console.WriteLine(transition.OuterXml);
        //             if (transition.AdvanceAfterTime.HasValue)
        //                 return transition.AdvanceAfterTime;
        //             break;
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         //Do nothing
        //     }

        //     return returnDuration;

        // }

        // public static int ConvertStringToInt(StringValue stringValue)
        // {
        //     int convertedInt = 0;
        //     try
        //     {
        //         Int32.TryParse(stringValue, out convertedInt);

        //     }
        //     catch (Exception)
        //     {

        //         throw;
        //     }
        //     return convertedInt;
        // }
    }
}
