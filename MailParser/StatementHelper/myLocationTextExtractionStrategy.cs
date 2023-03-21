using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using System.Collections;
using System.Collections.Generic;
using System;
using System.Text;
using iTextSharp.text;
using System.Linq;
using System.Diagnostics;
using System.Threading;
using Logger;

namespace PdfHelper
{
    internal class myLocationTextExtractionStrategy : ITextExtractionStrategy
    {
        public float UndercontentCharacterSpacing
        {
            get { return _undercontentCharacterSpacing; }
            set { _undercontentCharacterSpacing = value; }
        }
        private float _undercontentCharacterSpacing;


        public float UndercontentHorizontalScaling
        {
            get { return _undercontentHorizontalScaling; }
            set { _undercontentHorizontalScaling = value; }
        }
        private float _undercontentHorizontalScaling;


        public SortedList<string, DocumentFont> ThisPdfDocFonts { get; private set; }
        public bool DUMP_STATE;

        private List<TextChunk> locationalResult = new List<TextChunk>();


        public myLocationTextExtractionStrategy()
        {
            ThisPdfDocFonts = new SortedList<string, DocumentFont>();
        }

        public void BeginTextBlock()
        {
        }

        public void EndTextBlock()
        {
        }
        private readonly int QR_CODE_LEN = 24;

        public float def_space_width = 3.6f;
        public Dictionary<int, List<TextChunk>> line_chunks = new Dictionary<int, List<TextChunk>>();
        public string GetResultantText()
        {
            if (DUMP_STATE)
            {
                DumpState();

            }
            locationalResult.Sort();

            TextChunk lastChunk = null;
            int line_num = 0;
            line_chunks = new Dictionary<int, List<TextChunk>>();
            foreach (TextChunk chunk in locationalResult)
            {
                string chunk_text = chunk.text;
                bool is_QR = (chunk_text.Count(x => x == '1' || x == '0') == QR_CODE_LEN);
                if (is_QR)
                    continue;

                bool new_line = (lastChunk == null || !chunk.SameLine(lastChunk));
                if (new_line)
                    line_chunks.Add(++line_num, new List<TextChunk>() { chunk });
                else
                    line_chunks[line_num].Add(chunk);
                lastChunk = chunk;
            }
            if (line_chunks.Count != line_num)
            {
                lastChunk = null;
            }
            var sb = new StringBuilder();
            for (int i = 1; i <= line_num; i++)
            {
                if (!line_chunks.ContainsKey(i))
                    break;
                List<TextChunk> chunks = line_chunks[i];
                string line_text = "";
                foreach (TextChunk chunk in chunks)
                {
                    float f = line_text.Length * def_space_width;
                    while (f < chunk.distParallelStart)
                    {
                        line_text += " ";
                        f += def_space_width/*chunk.charSpaceWidth*/;
                    }
                    line_text += chunk.text;
                }
                sb.Append('\n');
                sb.Append(line_text);
            }
            string s = sb.ToString();
            return s;

            //sb = new StringBuilder();
            //lastChunk = null;
            //foreach (TextChunk chunk in locationalResult)
            //{
            //    bool new_line = (lastChunk == null || !chunk.SameLine(lastChunk));

            //    if (new_line)
            //    {
            //        if (lastChunk != null)
            //            sb.Append('\n');

            //        float dist = chunk.distParallelStart;
            //        if (dist < -chunk.charSpaceWidth)
            //        {
            //            sb.Append(' ');
            //        }
            //        else if (dist > chunk.charSpaceWidth / 2.0f)
            //        {
            //            float f = 0;
            //            while (f < dist)
            //            {
            //                sb.Append(' ');
            //                f += def_space_width/*chunk.charSpaceWidth*/;
            //            }
            //            if (!StartsWithSpace(chunk.text) && (lastChunk != null && !EndWithSpace(lastChunk.text)))
            //                sb.Append(' ');
            //        }
            //        sb.Append(chunk.text);
            //    }
            //    else
            //    {
            //        float dist = chunk.DistanceFromEndOf(lastChunk);
            //        //if (dist < -chunk.charSpaceWidth)
            //        //{
            //        //    sb.Append(' ');
            //        //}
            //        //else if (dist > chunk.charSpaceWidth / 2.0f && !StartsWithSpace(chunk.text) && !EndWithSpace(lastChunk.text))
            //        //{
            //        //    sb.Append(' ');
            //        //}

            //        if (dist < -chunk.charSpaceWidth)
            //        {
            //            sb.Append(' ');
            //        }
            //        else if (dist > chunk.charSpaceWidth / 2.0f)
            //        {
            //            float f = 0;
            //            while (f < dist)
            //            {
            //                sb.Append(' ');
            //                f += def_space_width/*chunk.charSpaceWidth*/;
            //            }
            //            if (!StartsWithSpace(chunk.text) && !EndWithSpace(lastChunk.text))
            //                sb.Append(' ');
            //        }
            //        sb.Append(chunk.text);
            //    }
            //    lastChunk = chunk;
            //}
            //return sb.ToString();
        }

        private bool StartsWithSpace(string str)
        {
            if (str.Length == 0)
            {
                return false;
            }
            return str[0] == ' ';
        }

        internal List<Rectangle> GetTextLocations(string pSearchString, StringComparison pStrComp)
        {
            var foundMatches = new List<iTextSharp.text.Rectangle>();
            var sb = new StringBuilder();
            var thisLineChunks = new List<TextChunk>();
            bool bStart = false, bEnd = false;
            TextChunk firstChunk = null;
            TextChunk lastChunk = null;
            string sTextInUsedChunks = "";

            //foreach (TextChunk chunk in locationalResult)
            for (int j = 0; j < locationalResult.Count; j++)
            {
                TextChunk chunk = locationalResult[j];
                if (chunk.text.Contains(pSearchString))
                    Thread.Sleep(1);
                if (thisLineChunks.Count > 0 && (!chunk.SameLine(thisLineChunks.Last()) || j == locationalResult.Count - 1)) //را اضافه کردlocationalResult.Countشرط Boris آقای
                {
                    if (sb.ToString().IndexOf(pSearchString, pStrComp) > -1)
                    {
                        string sLine = sb.ToString();

                        int iCount = 0;
                        int lPos;
                        lPos = sLine.IndexOf(pSearchString, 0, pStrComp);
                        while (lPos > -1)
                        {
                            iCount++;
                            if (lPos + pSearchString.Length > sLine.Length)
                                break;
                            else
                                lPos += pSearchString.Length;
                            lPos = sLine.IndexOf(pSearchString, lPos, pStrComp);
                        }
                        int curPos = 0;
                        for (int i = 1; i <= iCount; i++)
                        {
                            string sCurrentText; int iFromChar; int iToChar;
                            iFromChar = sLine.IndexOf(pSearchString, curPos, pStrComp);
                            curPos = iFromChar;
                            iToChar = iFromChar + pSearchString.Length - 1;
                            sCurrentText = "";
                            sTextInUsedChunks = "";
                            firstChunk = null;
                            lastChunk = null;

                            foreach (TextChunk chk in thisLineChunks)
                            {
                                sCurrentText = string.Concat(sCurrentText, chk.text);
                                if (!bStart && sCurrentText.Length - 1 >= iFromChar)
                                {
                                    firstChunk = chk;
                                    bStart = true;
                                }

                                if (bStart && !bEnd)
                                    //sCurrentText = string.Concat(sCurrentText, chk.text);
                                    sTextInUsedChunks = sTextInUsedChunks + chk.text;

                                if (!bEnd && sCurrentText.Length - 1 >= iToChar)
                                {
                                    lastChunk = chk;
                                    bEnd = true;
                                }

                                if (bStart && bEnd)
                                {
                                    foundMatches.Add(GetRectangleFromText(firstChunk, lastChunk, pSearchString, sTextInUsedChunks, iFromChar, iToChar, pStrComp));
                                    curPos += pSearchString.Length;
                                    bStart = bEnd = false;
                                    break;
                                }
                            }
                        }
                    }
                    sb.Clear();
                    thisLineChunks.Clear();
                }
                thisLineChunks.Add(chunk);
                sb.Append(chunk.text);
            }
            return foundMatches;
        }

        private Rectangle GetRectangleFromText(TextChunk firstChunk, TextChunk lastChunk, string pSearchString, string sTextinChunks, int iFromChar, int iToChar, StringComparison pStrComp)
        {
            float lineRealWidth = lastChunk.PosRight - firstChunk.PosLeft;

            float lineTextWidth = GetStringWidth(sTextinChunks, lastChunk.curFontSize, lastChunk.charSpaceWidth, ThisPdfDocFonts.Values.ElementAt(lastChunk.FontIndex)); // textطول متن در واحد

            float TransformationValue = lineRealWidth / lineTextWidth;

            int iStart = sTextinChunks.IndexOf(pSearchString, pStrComp);

            int iEnd = iStart + pSearchString.Length - 1;

            string sLeft;
            if (iStart == 0)
            {
                sLeft = null;
            }
            else
            {
                sLeft = sTextinChunks.Substring(0, iStart);
            }
            string sRight;
            if (iEnd == sTextinChunks.Length - 1)
            {
                sRight = null;
            }
            else
            {
                sRight = sTextinChunks.Substring(iEnd + 1, sTextinChunks.Length - iEnd - 1);
            }

            float leftWidth = 0;
            if (iStart > 0)
            {
                leftWidth = GetStringWidth(sLeft, lastChunk.curFontSize, lastChunk.charSpaceWidth, ThisPdfDocFonts.Values.ElementAt(lastChunk.FontIndex));
                leftWidth *= TransformationValue;
            }

            float rightWidth = 0;
            if (iEnd < sTextinChunks.Length - 1)
            {
                rightWidth = GetStringWidth(sRight, lastChunk.curFontSize, lastChunk.charSpaceWidth, ThisPdfDocFonts.Values.ElementAt(lastChunk.FontIndex));
                rightWidth *= TransformationValue;
            }
            float leftOffset = firstChunk.distParallelStart + leftWidth;
            float rightOffset = lastChunk.distParallelEnd - rightWidth;

            return new Rectangle(leftOffset, firstChunk.PosBottom, rightOffset, firstChunk.PosTop);
        }

        private float GetStringWidth(string str, float curFontSize, float pSingleSpaceWidth, DocumentFont pFont)
        {
            char[] chars = str.ToCharArray();
            float totalWidth = 0;
            float w = 0;
            foreach (char c in chars)
            {
                w = pFont.GetWidth(c) / 1000;
                totalWidth += (w * curFontSize + UndercontentCharacterSpacing) * UndercontentHorizontalScaling / 100;
            }
            return totalWidth;
        }


        private void DumpState()
        {
            foreach (TextChunk location in locationalResult)
            {
                location.PrintDiagnostics();
                MyLogger.Info("");
            }
        }

        public void RenderImage(ImageRenderInfo renderInfo)
        {
        }

        public void RenderText(TextRenderInfo renderInfo)
        {
            LineSegment segment = renderInfo.GetBaseline();
            var location = new TextChunk(renderInfo.GetText(), segment.GetStartPoint(), segment.GetEndPoint(), renderInfo.GetSingleSpaceWidth());
            Debug.Print(renderInfo.GetText());
            Vector horizonCoordinate = renderInfo.GetDescentLine().GetStartPoint();
            Vector verticalRight = renderInfo.GetAscentLine().GetEndPoint();
            location.PosLeft = horizonCoordinate[Vector.I1];
            location.PosRight = verticalRight[Vector.I1];
            location.PosBottom = horizonCoordinate[Vector.I2];
            location.PosTop = verticalRight[Vector.I2];
            location.curFontSize = location.PosTop - segment.GetStartPoint()[Vector.I2];
            string strKey = string.Concat(renderInfo.GetFont().PostscriptFontName, location.curFontSize);
            if (!ThisPdfDocFonts.ContainsKey(strKey))
            {
                ThisPdfDocFonts.Add(strKey, renderInfo.GetFont());
            }
            location.FontIndex = ThisPdfDocFonts.IndexOfKey(strKey);
            locationalResult.Add(location);
        }

        private bool EndWithSpace(string str)
        {
            if (str.Length == 0)
                return false;
            return str[str.Length - 1] == ' ';
        }
        public class TextChunk : IComparable<TextChunk>
        {
            public TextChunk(string str, Vector startLocation, Vector endLocation, float charSpaceWidth)
            {
                this.text = str;
                this.startLocation = startLocation;
                this.endLocation = endLocation;
                this.charSpaceWidth = charSpaceWidth;
                Vector oVector = endLocation.Subtract(startLocation);
                if (oVector.Length == 0)
                {
                    oVector = new Vector(1, 0, 0);
                }
                orientationVector = oVector.Normalize();
                orientationMagnitude = (int)Math.Truncate(Math.Atan2(orientationVector[Vector.I2], orientationVector[Vector.I1]) * 1000);
                var origin = new Vector(0, 0, 1);
                distPerpendicular = (int)((startLocation.Subtract(origin)).Cross(orientationVector)[Vector.I3]);
                distParallelStart = orientationVector.Dot(startLocation);
                distParallelEnd = orientationVector.Dot(endLocation);
            }

            internal float charSpaceWidth;
            internal float curFontSize { get; set; }
            internal float distParallelEnd;
            internal float distParallelStart;
            internal string text;
            private Vector startLocation;
            private Vector endLocation;
            private Vector orientationVector;
            private int orientationMagnitude;
            private int distPerpendicular;

            public int FontIndex { get; internal set; }
            public float PosBottom { get; internal set; }
            public float PosLeft { get; internal set; }
            public float PosRight { get; internal set; }
            public float PosTop { get; internal set; }

            public int CompareTo(TextChunk rhs)
            {
                if (this.Equals(rhs))
                {
                    return 0;
                }
                int rslt;
                rslt = CompareInts(orientationMagnitude, rhs.orientationMagnitude);
                if (rslt != 0)
                {
                    return rslt;
                }
                rslt = CompareInts(distPerpendicular, rhs.distPerpendicular);
                if (rslt != 0)
                {
                    return rslt;
                }
                rslt = distParallelStart < rhs.distParallelStart ? -1 : 1;
                return rslt;
            }

            private int CompareInts(int int1, int int2)
            {
                return (int1 == int2 ? 0 : (int1 < int2 ? -1 : 1)) ;
            }
            internal float DistanceFromEndOf(TextChunk other)
            {
                float distance = distParallelStart - other.distParallelEnd;
                return distance;
            }

            internal void PrintDiagnostics()
            {
                MyLogger.Info("Text (@" + Convert.ToString(startLocation) + " -> " + Convert.ToString(endLocation) + "): " + text);
                MyLogger.Info("orientationMagnitude: " + orientationMagnitude);
                MyLogger.Info("distPerpendicular: " + distPerpendicular);
                MyLogger.Info("distParallelStart: " + distParallelStart);
                MyLogger.Info("distParallelEnd: " + distParallelEnd);
            }

            internal bool SameLine(TextChunk a)
            {
                if (orientationMagnitude != a.orientationMagnitude)
                    return false;
                if (distPerpendicular != a.distPerpendicular)
                    return false;

                return true;
            }
        }
    }


}