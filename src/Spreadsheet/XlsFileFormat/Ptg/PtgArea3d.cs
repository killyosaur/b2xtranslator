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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgArea3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgArea3d;

        public UInt16 ixti;

        public UInt16 rwFirst;
        public UInt16 rwLast;
        public UInt16 colFirst;
        public UInt16 colLast;

        public bool rwFirstRelative;
        public bool rwLastRelative;
        public bool colFirstRelative;
        public bool colLastRelative;

        public PtgArea3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 11;
            this.ixti = this.Reader.ReadUInt16();
            this.rwFirst = this.Reader.ReadUInt16();
            this.rwLast = this.Reader.ReadUInt16();
            this.colFirst = this.Reader.ReadUInt16();
            this.colLast = this.Reader.ReadUInt16();

            this.colFirstRelative = Utils.BitmaskToBool(colFirst, 0x4000);
            this.rwFirstRelative = Utils.BitmaskToBool(colFirst, 0x8000);
            this.colLastRelative = Utils.BitmaskToBool(colLast, 0x4000);
            this.rwLastRelative = Utils.BitmaskToBool(colLast, 0x8000);

            colFirst = (UInt16)(colFirst & 0x3FFF);
            colLast = (UInt16)(colLast & 0x3FFF);

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}


