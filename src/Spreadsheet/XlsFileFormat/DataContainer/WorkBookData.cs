/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;



namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class is a container for the extracted data 
    /// Implements the IVisitable Interface 
    /// </summary>
    public class WorkBookData : IVisitable
    {
        /// <summary>
        /// This attribute stores the SharedStringTable Data
        /// </summary>
        private SSTData sstData;
        public SSTData SstData
        {
            get { return sstData; }
            set { sstData = value; }
        }


        public bool Template { get; set; }

        public List<SheetData> boundSheetDataList;
        public List<ExternSheetData> externSheetDataList;
        public LinkedList<SupBookData> supBookDataList;
        public LinkedList<XTIData> xtiDataList;
        public List<Lbl> definedNameList; 

        public int refWorkBookNumber;

        public StyleData.StyleData styleData;

        public BOF BOF;
        public CodeName CodeName;

        /// <summary>
        /// Ctor 
        /// </summary>
        public WorkBookData()
        {
            this.boundSheetDataList = new List<SheetData>();
            this.externSheetDataList = new List<ExternSheetData>();
            this.supBookDataList = new LinkedList<SupBookData>();
            this.xtiDataList = new LinkedList<XTIData>();
            this.definedNameList = new List<Lbl>(); 
            refWorkBookNumber = 0;

            this.styleData = new StyleData.StyleData(); 
        }

        /// <summary>
        /// Adds a WorkSheetData Element to the internal list 
        /// </summary>
        /// <param name="bsd">The Boundsheetdata element</param>
        public void addBoundSheetData(SheetData bsd)
        {
            this.boundSheetDataList.Add(bsd); 
        }

        /// <summary>
        /// Add the ExternSheetData extracted from an EXTERNSHEET BIFF Record 
        /// </summary>
        /// <param name="ext">BIFF Record</param>
        public void addExternSheetData(ExternSheet ext)
        {
            for (int i = 0; i < ext.cXTI; i++)
            {
                ExternSheetData extdata = new ExternSheetData(ext.iSUPBOOK[i], ext.itabFirst[i], ext.itabLast[i]);
                this.externSheetDataList.Add(extdata); 
            }
        }

        /// <summary>
        /// Add a SUPBOOK BIFF Record to the list 
        /// </summary>
        /// <param name="sup"></param>
        public void addSupBookData(SupBook sup)
        {
            
            SupBookData supbook = new SupBookData(sup);
            if (!supbook.SelfRef)
            {
                this.refWorkBookNumber++;
                supbook.Number = this.refWorkBookNumber; 
            }


            this.supBookDataList.AddLast(supbook); 
        }


        /// <summary>
        /// Returns the string from an external workbook 
        /// </summary>
        /// <param name="index">index from the workbook </param>
        /// <returns>Filename and sheetname </returns>
        public String getIXTIString(UInt16 index)
        {
            ExternSheetData extSheet = this.externSheetDataList[index];
            SupBookData supData = null;
            LinkedList<SupBookData>.Enumerator listenum = this.supBookDataList.GetEnumerator();   
            

            int count = 0;
            listenum.MoveNext(); 
            do
            {
                if (count == extSheet.iSUPBOOK)
                {
                    supData = listenum.Current; 
                }
                count++; 
            }
            while (listenum.MoveNext());

            String back = ""; 
            if (supData != null && supData.SelfRef)
            {
                String first = this.boundSheetDataList[extSheet.itabFirst].boundsheetRecord.stName.Value;
                String last = this.boundSheetDataList[extSheet.itabLast].boundsheetRecord.stName.Value;
                if (first.Equals(last))
                {
                    back = first;
                }
                else
                {
                    back = first + ":" + last; 
                }
            }
            else
            {
                String first = supData.getRgstString(extSheet.itabFirst);
                String last = supData.getRgstString(extSheet.itabLast);
                if (first.Equals(last))
                {
                    back = first;
                }
                else
                {
                    back = first + ":" + last;
                }
                // add one to the index 
                back = "[" + supData.Number.ToString()+"]" + back; 

            }
            return back; 
        }


        /// <summary>
        /// Add a XCT Data structure to the internal stack 
        /// </summary>
        /// <param name="xct"></param>
        public void addXCT(XCT xct)
        {
            XTIData xti = new XTIData(this.xtiDataList.Count - 1, this.supBookDataList.Count - 1, xct.itab);
            this.xtiDataList.AddLast(xti);
            this.supBookDataList.Last.Value.addXCT(xct); 
        }

        /// <summary>
        /// Add a CRN Data structure to the internal list 
        /// </summary>
        /// <param name="xct"></param>
        public void addCRN(CRN crn)
        {
            this.supBookDataList.Last.Value.addCRN(crn); 
        }

        /// <summary>
        /// Add a EXTERNNAME Data structure to the internal list 
        /// </summary>
        /// <param name="xct"></param>
        public void addEXTERNNAME(ExternName extname)
        {
            this.supBookDataList.Last.Value.addEXTERNNAME(extname); 
        }

        /// <summary>
        /// add a definedName data object
        /// </summary>
        /// <param name="name"></param>
        public void addDefinedName(Lbl name)
        {
            this.definedNameList.Add(name); 
        }

        /// <summary>
        /// Get the definedname string from an ID 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public String getDefinedNameByRef(int id)
        {
            if (this.definedNameList[id - 1].Name.Value.Length > 1)
            {
                return this.definedNameList[id - 1].Name.Value; 
            }
            else
            {
                string internName = "_xlnm." + ExcelHelperClass.getNameStringfromBuiltInFunctionID(this.definedNameList[id - 1].Name.Value);
                return internName; 
            }
           
        }

        /// <summary>
        /// returns the extern name if an valid ID is given! 
        /// </summary>
        /// <param name="supIndex"></param>
        /// <param name="nameIndex"></param>
        /// <returns></returns>
        public String getExternNameByRef(UInt16 supIndex, UInt32 nameIndex)
        {
            ExternSheetData extSheet = this.externSheetDataList[supIndex];
            SupBookData supData = null;
            LinkedList<SupBookData>.Enumerator listenum = this.supBookDataList.GetEnumerator();

            String back = ""; 
            int count = 0;
            int counttwo = 0; 
            listenum.MoveNext();
            do
            {
                if (count == extSheet.iSUPBOOK)
                {
                    supData = listenum.Current;
                }
                count++;
            }
            while (listenum.MoveNext());

            LinkedList<String>.Enumerator nameEnum = supData.ExternNames.GetEnumerator();
            do
            {
                if (counttwo == nameIndex)
                {
                    back = nameEnum.Current;
                }
                counttwo++;
            } while (nameEnum.MoveNext());
            

            return "[" + (supData.Number) + "]!" + back; 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WorkBookData>)mapping).Apply(this);
        }

        #endregion


    }
}
