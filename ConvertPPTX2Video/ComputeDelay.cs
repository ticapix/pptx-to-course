using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace PPTX2Course
{

    public class PPTXComputeDelay {


        public static int GetSlideTransitionsDuration(Slide slide)
        {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-2.8.1
            var transitions = slide.Descendants<Transition>();
            foreach (var transition in transitions) {
                if (transition.Duration != null && transition.Duration.HasValue) {
                        return ConvertStringToInt(transition.Duration);
                }
            }
            return 0; // TODO add default value
        }

        public static int GetSlideAnimationsDuration(Slide slide)
        {
            return ComputeTimingDelay(slide.Timing);
        }

        public static int GetSlideAdvanceAfterTimeDuration(Slide slide)
        {
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-2.8.1
            var transitions = slide.Descendants<Transition>();
            foreach (var transition in transitions) {
                if (transition.AdvanceAfterTime != null && transition.AdvanceAfterTime.HasValue) {
                        return ConvertStringToInt(transition.AdvanceAfterTime);
                }
            }
            return 0; // TODO add default value
        }

        public static int ConvertStringToInt(StringValue stringValue)
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
        private static int ComputeTimingDelay(Timing node) {
            int delay = 0;
            // http://www.datypic.com/sc/ooxml/e-p_tnLst-1.html
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timenodelist?view=openxml-2.8.1
            var tmRoot = node.TimeNodeList.ParallelTimeNode.CommonTimeNode;
            // http://www.datypic.com/sc/ooxml/a-nodeType-1.html
            Debug.Assert(tmRoot.NodeType == TimeNodeValues.TmingRoot, "the found timeNode is not the root one");

            delay += ComputeParallelTimeNodeDelay(node.TimeNodeList.ParallelTimeNode);
            return delay;
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
        private static int ComputeTimeNodeListDelay(TimeTypeListType nodeList) {
            int delay = 0;
            foreach (var node in nodeList) {
                if (node is ParallelTimeNode) { // http://www.datypic.com/sc/ooxml/e-p_par-1.html
                    delay += ComputeParallelTimeNodeDelay(node as ParallelTimeNode);
                }
                else if (node is SequenceTimeNode) { // http://www.datypic.com/sc/ooxml/e-p_seq-1.html
                    delay += ComputeSequenceTimeNodeDelay(node as SequenceTimeNode);
                }
                else if (node is SetBehavior) { // http://www.datypic.com/sc/ooxml/e-p_set-1.html
                    delay += ComputeCommonTimeNodeDelay((node as SetBehavior).CommonBehavior.CommonTimeNode);
                }
                else if (node is AnimateEffect) { // http://www.datypic.com/sc/ooxml/e-p_animEffect-1.html
                    delay += ComputeCommonTimeNodeDelay((node as AnimateEffect).CommonBehavior.CommonTimeNode);
                } else {
                    Console.WriteLine($"unknown node: {node.OuterXml}");
                }
            }
            return delay;
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_seq-1.html
        Sequence [1..1]
            p:cTn [1..1]    Common Time Node Properties
            p:prevCondLst [0..1]    Previous Conditions List
            p:nextCondLst [0..1]    Next Conditions List
        */
        private static int ComputeSequenceTimeNodeDelay(SequenceTimeNode node) {
            int delay = 0;

            delay += ComputeCommonTimeNodeDelay(node.CommonTimeNode);
            if (node.PreviousConditionList != null && node.PreviousConditionList.HasChildren) {
                delay += ComputeTimeConditionListDelay(node.PreviousConditionList);
            }
            if (node.NextConditionList != null && node.NextConditionList.HasChildren) {
                delay += ComputeTimeConditionListDelay(node.NextConditionList);
            }
            return delay;
        }

        /*
        http://www.datypic.com/sc/ooxml/e-p_par-1.html
        p:cTn [1..1]    Parallel TimeNode
        */
        private static int ComputeParallelTimeNodeDelay(ParallelTimeNode node) {
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
        private static int ComputeCommonTimeNodeDelay(CommonTimeNode node) {
            int delay = 0;
            if (node.Duration != null && node.Duration.HasValue) {
                delay += ConvertStringToInt(node.Duration);
            }
            if (node.StartConditionList != null && node.StartConditionList.HasChildren) {
                delay += ComputeTimeConditionListDelay(node.StartConditionList);
            }
            if (node.EndConditionList != null && node.EndConditionList.HasChildren) {
            delay += ComputeTimeConditionListDelay(node.EndConditionList);
            }
            if (node.ChildTimeNodeList != null && node.ChildTimeNodeList.HasChildren) {
                delay += ComputeTimeNodeListDelay(node.ChildTimeNodeList as TimeTypeListType);
            }
            int repeatCount = 1;
            if (node.RepeatCount != null && node.RepeatCount.HasValue) {
                repeatCount = ConvertStringToInt(node.RepeatCount) / 1000;
            }
            return repeatCount * delay;
        }


        /*
        http://www.datypic.com/sc/ooxml/e-p_stCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_endCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_prevCondLst-1.html
        http://www.datypic.com/sc/ooxml/e-p_nextCondLst-1.html
        p:cond [1..*]    Condition
        */
        private static int ComputeTimeConditionListDelay(TimeListTimeConditionalListType nodeList) {
            int delay = 0;
            foreach (Condition cond in nodeList.Cast<Condition>()) {
                if (cond.Delay != null && cond.Delay.HasValue) {
                    delay += ConvertStringToInt(cond.Delay);
                }
            }
            return delay;
        }

    }
}