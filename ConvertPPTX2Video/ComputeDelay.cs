using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using FFMpegCore.Arguments;

// using Microsoft.Extensions.Logging;

// using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
// ILogger logger = factory.CreateLogger("Program");
// logger.LogInformation("Hello World! Logging is {Description}.", "fun");

namespace PPTX2Course
{

    public class PPTXComputeDelay(Slide slide, int defaultTransitionDurationMs)
    {
        private Slide _slide = slide;
        private int _defaultTransitionDurationMs = defaultTransitionDurationMs;
        private int _level = 1; // used to pretty print when browsing the XML

        public int GetSlideTransitionsDuration()
        {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-2.8.1
            var transitions = _slide.Descendants<Transition>();
            foreach (var transition in transitions)
            {
                if (transition.Duration != null && transition.Duration.HasValue)
                {
                    return ConvertStringToInt(transition.Duration);
                }
            }
            return 0;
        }

        public int GetSlideAnimationsDuration()
        {
            if (_slide.Timing != null)
            {
                return ComputeTimingDelay(_slide.Timing);
            }
            return 0;
        }

        public int GetSlideAdvanceAfterTimeDuration()
        {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-2.8.1
            var transitions = _slide.Descendants<Transition>();
            foreach (var transition in transitions)
            {
                if (transition.AdvanceAfterTime != null && transition.AdvanceAfterTime.HasValue)
                {
                    return ConvertStringToInt(transition.AdvanceAfterTime);
                }
            }
            return _defaultTransitionDurationMs;
        }

        private static int ConvertStringToInt(StringValue stringValue)
        {
            int convertedInt;
            return Int32.TryParse(stringValue, out convertedInt) ? convertedInt : 0;
        }

        /*
        http://www.datypic.com/sc/ooxml/t-p_CT_SlideTiming.html
        Sequence [1..1]
            p:tnLst [0..1]    Time Node List
            p:bldLst [0..1]    Build List
            p:extLst [0..1]    Extension List
        */
        private int ComputeTimingDelay(Timing node)
        {
            // http://www.datypic.com/sc/ooxml/t-p_CT_SlideTiming.html
            CommonTimeNode tmRoot = node.TimeNodeList.ParallelTimeNode.CommonTimeNode;
            // http://www.datypic.com/sc/ooxml/a-nodeType-1.html
            Debug.Assert(tmRoot.NodeType == TimeNodeValues.TmingRoot, "The found timeNode is not of type tmRoot.");
            Debug.Assert(tmRoot.StartConditionList == null);
            Debug.Assert(tmRoot.EndConditionList == null);
            Debug.Assert(tmRoot.ChildTimeNodeList.Count() == 1, "There is more than one item under tmRoot.");
            OpenXmlElement elt = tmRoot.ChildTimeNodeList.First();
            Debug.Assert(elt is SequenceTimeNode, "The found element is not of class SequenceTimeNode.");
            SequenceTimeNode mainSeq = elt as SequenceTimeNode;
            Debug.Assert(mainSeq.CommonTimeNode.NodeType == TimeNodeValues.MainSequence, "The found timeNode is not of type mainSeq.");
            ChildTimeNodeList animationList = mainSeq.CommonTimeNode.ChildTimeNodeList;

            Debug.Assert(animationList.Count() == animationList.Elements<ParallelTimeNode>().Count()); // we should only have ParallelTimeNode elements
            Console.WriteLine($"Found {animationList.Count()} animation(s) in the slide.");
            List<int> parallelDelays = [];
            foreach (ParallelTimeNode nodeAnimation in animationList.Elements<ParallelTimeNode>())
            {
                // Console.WriteLine(nodeAnimation.OuterXml);
                parallelDelays.Add(ComputeParallelTimeNodeDelay(nodeAnimation));
                Console.WriteLine($"Animation duration: {parallelDelays.Last()} ms");
            }
            return parallelDelays.Sum(); // !! WARNING: no idea why MainSeq is composed of ParallelTimeNode and not SequenceTimeNode as one would expect.
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_childTnLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_subTnLst-1.html
        Choice [1..*]
            p:par    Parallel Time Node
            p:seq    Sequence Time Node
            p:excl    Exclusive
            p:anim    Animate
            p:animClr    Animate Color Behavior
            p:animEffect    Animate Effect
            p:animMotion    Animate Motion
            p:animRot    Animate Rotation
            p:animScale    Animate Scale
            p:cmd    Command
            p:set    Set Time Node Behavior
            p:audio    Audio
            p:video    Video
        */
        private int ComputeTimeNodeListDelay(TimeTypeListType nodeList)
        {
            _level += 1;
            { // expect only ParallelTimeNode
                List<int> parallelDelays = [];
                foreach (ParallelTimeNode node in nodeList.Elements<ParallelTimeNode>())
                {
                    parallelDelays.Add(ComputeParallelTimeNodeDelay(node));
                }
                if (parallelDelays.Count != 0)
                {
                    _level += -1;
                    Debug.Assert(parallelDelays.Count == nodeList.Count());
                    IEnumerable<string> nodeNameDurations = nodeList.Zip(parallelDelays, (node, delay) => $"{node.LocalName} ({delay} ms)");
                    Console.WriteLine(new string('|', _level) + $"found {nodeNameDurations.Count()} elements in the list with types: " + string.Join(", ", nodeNameDurations));
                    return parallelDelays.Max();
                }
            }
            { // expect only SequenceTimeNode
                List<int> sequenceDelays = [];
                foreach (SequenceTimeNode node in nodeList.Elements<SequenceTimeNode>())
                {
                    sequenceDelays.Add(ComputeSequenceTimeNodeDelay(node));
                }
                if (sequenceDelays.Count != 0)
                {
                    _level += -1;
                    Debug.Assert(sequenceDelays.Count == nodeList.Count());
                    IEnumerable<string> nodeNameDurations = nodeList.Zip(sequenceDelays, (node, delay) => $"{node.LocalName} ({delay} ms)");
                    Console.WriteLine(new string('|', _level) + $"found {nodeNameDurations.Count()} elements in the list with types: " + string.Join(", ", nodeNameDurations));
                    return sequenceDelays.Sum();
                }
            }
            { // expect somethings only different from ParallelTimeNode and SequenceTimeNode
                int delay = 0;
                IEnumerable<string> nodeNames = nodeList.Select(node => node.LocalName);
                Console.WriteLine(new string('|', _level) + $"found {nodeNames.Count()} elements in the list with types: " + string.Join(", ", nodeNames));
                foreach (OpenXmlElement node in nodeList)
                {
                    Debug.Assert(node is not ParallelTimeNode);
                    Debug.Assert(node is not SequenceTimeNode);
                    if (node is SetBehavior)
                    { // http://www.datypic.com/sc/ooxml/e-p_set-1.html
                        delay += ComputeCommonTimeNodeDelay((node as SetBehavior)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is AnimateColor)
                    { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                        // in our tests, setBehavior seems to already contain the duration of the effect.
                        // skipping the node as it artificially increase animation duration
                        // delay += ComputeCommonTimeNodeDelay((node as AnimateColor)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is AnimateEffect)
                    { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                        delay += ComputeCommonTimeNodeDelay((node as AnimateEffect)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is AnimateMotion)
                    { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                        delay += ComputeCommonTimeNodeDelay((node as AnimateMotion)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is AnimateRotation)
                    { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                        delay += ComputeCommonTimeNodeDelay((node as AnimateRotation)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is AnimateScale)
                    { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                        delay += ComputeCommonTimeNodeDelay((node as AnimateScale)!.CommonBehavior!.CommonTimeNode);
                    }
                    else if (node is Animate)
                    {
                        // Do nothing
                    }
                    else
                    {
                        Console.WriteLine($"unknown node: {node.OuterXml}");
                    }
                }
                _level += -1;
                // IEnumerable<string> nodeNames = nodeList.Select(node => node.LocalName);
                Console.WriteLine(new string('|', _level) + $"found {nodeNames.Count()} elements in the list with types: " + string.Join(", ", nodeNames) + $"({delay} ms)");
                return delay;
            }
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_seq-1.html
        Sequence [1..1]
            p:cTn [1..1]    Common Time Node Properties
            p:prevCondLst [0..1]    Previous Conditions List
            p:nextCondLst [0..1]    Next Conditions List
        */
        private int ComputeSequenceTimeNodeDelay(SequenceTimeNode node)
        {
            int delay = 0;

            delay += ComputeCommonTimeNodeDelay(node!.CommonTimeNode);
            if (node.PreviousConditionList != null && node.PreviousConditionList.HasChildren)
            {
                delay += ComputeTimeConditionListDelay(node.PreviousConditionList);
            }
            if (node.NextConditionList != null && node.NextConditionList.HasChildren)
            {
                delay += ComputeTimeConditionListDelay(node.NextConditionList);
            }
            return delay;
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_par-1.html
        p:cTn [1..1]    Parallel TimeNode
        */
        private int ComputeParallelTimeNodeDelay(ParallelTimeNode node)
        {
            return ComputeCommonTimeNodeDelay(node.CommonTimeNode);
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_cTn-1.html
        Sequence [1..1]
            p:stCondLst [0..1]    Start Conditions List
            p:endCondLst [0..1]    End Conditions List
            p:endSync [0..1]    EndSync
            p:iterate [0..1]    Iterate
            p:childTnLst [0..1]    Children Time Node List
            p:subTnLst [0..1]    Sub-TimeNodes List
        */
        private int ComputeCommonTimeNodeDelay(CommonTimeNode node)
        {
            int nodeDuration = 0;
            int startCondDelay = 0;
            int endCondDelay = 0;
            int delay = 0;
            if (node.Duration != null && node.Duration.HasValue)
            {
                nodeDuration = ConvertStringToInt(node.Duration);
            }
            if (node.StartConditionList != null && node.StartConditionList.HasChildren)
            {
                startCondDelay = ComputeTimeConditionListDelay(node.StartConditionList);
            }
            if (node.EndConditionList != null && node.EndConditionList.HasChildren)
            {
                endCondDelay = ComputeTimeConditionListDelay(node.EndConditionList);
            }
            if (node.ChildTimeNodeList != null && node.ChildTimeNodeList.HasChildren)
            {
                delay = ComputeTimeNodeListDelay(node.ChildTimeNodeList as TimeTypeListType);
            }
            int repeatCount = 1;
            if (node.RepeatCount != null && node.RepeatCount.HasValue)
            {
                repeatCount = ConvertStringToInt(node.RepeatCount) / 1000;
            }
            if (nodeDuration != 0)
            {
                Debug.Assert(delay == 0);
            }
            if (delay != 0) {
                Debug.Assert(nodeDuration == 0);
            }
            return startCondDelay + (repeatCount * (nodeDuration + delay)) + endCondDelay;
        }


        /*
        http://www.datypic.com/sc/ooxml/e-p_stCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_endCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_prevCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_nextCondLst-1.html
        p:cond [1..*]    Condition
        */
        private static int ComputeTimeConditionListDelay(TimeListTimeConditionalListType nodeList)
        {
            int delay = 0;
            foreach (Condition cond in nodeList.Cast<Condition>())
            {
                if (cond.Delay != null && cond.Delay.HasValue)
                {
                    delay += ConvertStringToInt(cond.Delay);
                }
            }
            return delay;
        }

    }
}