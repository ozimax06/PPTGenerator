using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace SlideMerge
{
    public class MergeHandler
    {
        [ThreadStatic]
        static uint uniqueId;

        [ThreadStatic]
        static int mergedSideCount = 0;

        [ThreadStatic]
        static int reOrderConstant = 0;

        [ThreadStatic]
        static int id = 0;        
                
        internal void MergeAllSlides(string parentPresentation, string destinationFileName, List<PresentationInfo> mergingFileList)
        {
            // key=>current slide position, value=>new position of the slide
            Dictionary<int, int> reOrderPair = new Dictionary<int, int>(); 
            

            parentPresentation = GetFileNameWithExtension(parentPresentation);
            destinationFileName = GetFileNameWithExtension(destinationFileName);
            string presentationTemplate = parentPresentation;

            if (presentationTemplate != destinationFileName)
                File.Copy(presentationTemplate, destinationFileName, true);

            mergedSideCount = GetTotalSlidesCount(parentPresentation);

            foreach (PresentationInfo importedFile in mergingFileList)
            {
                try
                {
                    int val = (importedFile.InsertPosition - 1) + reOrderPair.Count;
                    this.MergeSlides(importedFile, destinationFileName, ref reOrderPair, ref val, out reOrderConstant);
                }
                catch
                {
                    continue; // try merging other presentation files
                }
            }

            // re-order the slides to the actual position                                        
            foreach (KeyValuePair<int, int> reOrderValue in reOrderPair)
            {
                try
                {
                    this.ReorderSlides(destinationFileName, reOrderValue.Key, reOrderValue.Value);                    
                }
                catch
                {
                    continue;
                }
            }
        }

        public int GetTotalSlidesCount(string presentationFile)
        {
            using (PresentationDocument prstDoc = PresentationDocument.Open(presentationFile, false))
            {
                int slideCount = prstDoc.PresentationPart.Presentation.SlideIdList.Count();
                prstDoc.Close();
                return slideCount;
            }
        }
                
        private string GetFileNameWithExtension(string fileName)
        {
            string directory = System.IO.Path.GetDirectoryName(fileName);
            int index = fileName.LastIndexOf('.');
            if (index != -1)
            {
                if (fileName.Substring(index, fileName.Length - index).Contains("pptx") == true)
                    fileName = fileName.Substring(0, index);
            }
            return System.IO.Path.Combine(directory, fileName + ".pptx");
        }
                
        private void MergeSlides(PresentationInfo sourcePresentation, string destPresentation, ref Dictionary<int, int> reOrderPair, ref int val, out int reOrderConstantValue)
        {
            using (PresentationDocument destinationPresentationDoc = PresentationDocument.Open(destPresentation, true))
            {
                PresentationPart destPresPart = destinationPresentationDoc.PresentationPart;

                // If the merged presentation doesn't have a SlideIdList element yet then add it.
                if (destPresPart.Presentation.SlideIdList == null)
                    destPresPart.Presentation.SlideIdList = new SlideIdList();

                string sourceFileName = GetFileNameWithExtension(sourcePresentation.File);

                // Open the source presentation. This will throw an exception if the source presentation does not exist.
                using (PresentationDocument sourcePresentationDoc = PresentationDocument.Open(sourceFileName, false))
                {
                    PresentationPart sourcePresPart = sourcePresentationDoc.PresentationPart;

                    // Get unique ids for the slide master and slide lists for use later.
                    uniqueId = GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);
                    uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);

                    reOrderConstantValue = sourcePresPart.Presentation.SlideIdList.Count();

                    // Copy each slide in the source presentation in order to the destination presentation.
                    foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList)
                    {
                        SlidePart sp;
                        SlidePart destSp;
                        SlideMasterPart destMasterPart;
                        string relId;
                        SlideMasterId newSlideMasterId;
                        SlideId newSlideId;

                        //increase the slide count
                        mergedSideCount++;

                        if (sourcePresentation.InsertPosition != -1)
                        {
                            reOrderPair.Add(mergedSideCount - 1, val++);
                        }

                        // Create a unique relationship id.
                        id++;
                        sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
                        relId = "uniq" + id;

                        // Add the slide part to the destination presentation.
                        destSp = destPresPart.AddPart<SlidePart>(sp, relId);

                        // The master part was added. Make sure the relationship is in place.
                        destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                        destPresPart.AddPart(destMasterPart);

                        // Add slide master to slide master list.
                        uniqueId++;
                        newSlideMasterId = new SlideMasterId();
                        newSlideMasterId.RelationshipId = destPresPart.GetIdOfPart(destMasterPart);
                        newSlideMasterId.Id = uniqueId;

                        // Add slide to slide list.
                        maxSlideId++;
                        newSlideId = new SlideId();
                        newSlideId.RelationshipId = relId;
                        newSlideId.Id = maxSlideId;

                        destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                        destPresPart.Presentation.SlideIdList.Append(newSlideId);

                    }

                    // Make sure all slide ids are unique.
                    ModifySlideLayoutIds(destPresPart);
                }

                // Save the changes to the destination presentation.
                destPresPart.Presentation.Save();
            }
        }
                
        private void ModifySlideLayoutIds(PresentationPart presPart)
        {
            // Make sure all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = (uint)uniqueId;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }
                
        private uint GetMaxSlideId(SlideIdList slideIdList)
        {
            // Slide identifiers have a minimum value of greater than or equal to 256
            // and a maximum value of less than 2147483648. 
            uint max = 256;

            if (slideIdList != null)
                // Get the maximum id value from the current set of children.
                foreach (SlideId child in slideIdList.Elements<SlideId>())
                {
                    uint id = child.Id;

                    if (id > max)
                        max = id;
                }

            return max;
        }
                
        private uint GetMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
        {
            // Slide master identifiers have a minimum value of greater than or equal to 2147483648. 
            uint max = 2147483648;

            if (slideMasterIdList != null)
                // Get the maximum id value from the current set of children.
                foreach (SlideMasterId child in slideMasterIdList.Elements<SlideMasterId>())
                {
                    uint id = child.Id;

                    if (id > max)
                        max = id;
                }

            return max;
        }
                
        private int ReorderSlides(string fileName, int currentSlidePosition, int newPosition)
        {
            int returnValue = -1;

            if (newPosition == currentSlidePosition)
            {
                return returnValue;
            }

            using (PresentationDocument doc = PresentationDocument.Open(fileName, true))
            {
                PresentationPart presentationPart = doc.PresentationPart;

                if (presentationPart == null)
                {
                    throw new ArgumentException("Presentation part not found.");
                }

                int slideCount = presentationPart.SlideParts.Count();

                if (slideCount == 0)
                {
                    return returnValue;
                }

                int maxPosition = slideCount - 1;

                CalculatePositions(ref currentSlidePosition, ref newPosition, maxPosition);

                if (newPosition != currentSlidePosition)
                {
                    DocumentFormat.OpenXml.Presentation.Presentation presentation = presentationPart.Presentation;
                    SlideIdList slideIdList = presentation.SlideIdList;

                    SlideId sourceSlide = (SlideId)(slideIdList.ChildElements[currentSlidePosition]);
                    SlideId targetSlide = (SlideId)(slideIdList.ChildElements[newPosition]);

                    sourceSlide.Remove();

                    if (newPosition > currentSlidePosition)
                    {
                        slideIdList.InsertAfter(sourceSlide, targetSlide);
                    }
                    else
                    {
                        slideIdList.InsertBefore(sourceSlide, targetSlide);
                    }

                    returnValue = newPosition;

                    presentation.Save();
                }
            }
            return returnValue;
        }
                
        private void CalculatePositions(ref int originalPosition, ref int newPosition, int maxPosition)
        {
            if (originalPosition < 0)
            {
                originalPosition = maxPosition;
            }

            if (newPosition < 0)
            {
                newPosition = maxPosition;
            }

            if (originalPosition > maxPosition)
            {
                originalPosition = maxPosition;
            }
            if (newPosition > maxPosition)
            {
                newPosition = maxPosition;
            }
        }

    }
}
