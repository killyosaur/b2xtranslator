/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class PropertiesMapping : AbstractOpenXmlMapping
    {
        public PropertiesMapping(XmlWriter writer)
            : base(writer)
        {
        }

        protected void appendFlagAttribute(XmlElement node, SinglePropertyModifier sprm, string attributeName)
        {
            XmlAttribute att = _nodeFactory.CreateAttribute("w", attributeName, OpenXmlNamespaces.WordprocessingML);
            att.Value = sprm.Arguments[0].ToString();
            node.Attributes.Append(att);
        }

        protected virtual void appendFlagElement(XmlElement node, SinglePropertyModifier sprm, string elementName, bool unique)
        {
            XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);
            if (sprm.Arguments[0] == 0)
            {
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                val.Value = "false";
                ele.Attributes.Append(val);
            }
            if (unique)
            {
                foreach (XmlElement exEle in node.ChildNodes)
                {
                    if (exEle.Name == ele.Name)
                    {
                        node.RemoveChild(exEle);
                        break;
                    }
                }
            }
            node.AppendChild(ele);
        }

        protected void appendValueAttribute(XmlElement node, string attributeName, string attributeValue)
        {
            appendValueAttribute(node, "w", attributeName, attributeValue, OpenXmlNamespaces.WordprocessingML);
        }

        protected void appendValueAttribute(XmlElement node, string prefix, string attributeName, string attributeValue, string ns)
        {
            XmlAttribute att = _nodeFactory.CreateAttribute(prefix, attributeName, ns);
            att.Value = attributeValue;
            node.Attributes.Append(att);
        }

        protected void appendValueElement(XmlElement node, string elementName, string elementValue, bool unique)
        {
            XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);

            if(elementValue != null && elementValue != "")
            {
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                val.Value = elementValue;
                ele.Attributes.Append(val);
            }
            if (unique)
            {
                foreach (XmlElement exEle in node.ChildNodes)
                {
                    if (exEle.Name == ele.Name)
                    {
                        node.RemoveChild(exEle);
                        break;
                    }
                }
            }
            node.AppendChild(ele);
        }

        protected void appendDxaElement(XmlElement node, string elementName, string elementValue, bool unique)
        {
            XmlElement ele = _nodeFactory.CreateElement("w", elementName, OpenXmlNamespaces.WordprocessingML);
            XmlAttribute val = _nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
            val.Value = elementValue;
            ele.Attributes.Append(val);
            XmlAttribute type = _nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
            type.Value = "dxa";
            ele.Attributes.Append(type);
            if (unique)
            {
                foreach (XmlElement exEle in node.ChildNodes)
                {
                    if (exEle.Name == ele.Name)
                    {
                        node.RemoveChild(exEle);
                        break;
                    }
                }
            }
            node.AppendChild(ele);
        }

        protected void addOrSetBorder(XmlNode pBdr, XmlNode border)
        {
            //remove old border if it exist
            foreach (XmlNode bdr in pBdr.ChildNodes)
            {
                if (bdr.Name == border.Name)
                {
                    pBdr.RemoveChild(bdr);
                    break;
                }
            }

            //add new
            pBdr.AppendChild(border);
        }

        protected void appendBorderAttributes(BorderCode brc, XmlNode border)
        {
            if (brc.fNil)
            {
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                val.Value = "nil";
                border.Attributes.Append(val);
            }
            else
            {
                XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
                val.Value = getBorderType(brc.brcType);
                border.Attributes.Append(val);

                XmlAttribute color = _nodeFactory.CreateAttribute("w", "color", OpenXmlNamespaces.WordprocessingML);
                color.Value = new RGBColor(brc.cv, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                border.Attributes.Append(color);

                XmlAttribute space = _nodeFactory.CreateAttribute("w", "space", OpenXmlNamespaces.WordprocessingML);
                space.Value = brc.dptSpace.ToString();
                border.Attributes.Append(space);

                XmlAttribute sz = _nodeFactory.CreateAttribute("w", "sz", OpenXmlNamespaces.WordprocessingML);
                sz.Value = brc.dptLineWidth.ToString();
                border.Attributes.Append(sz);

                if (brc.fShadow)
                {
                    XmlAttribute shadow = _nodeFactory.CreateAttribute("w", "shadow", OpenXmlNamespaces.WordprocessingML);
                    shadow.Value = "1";
                    border.Attributes.Append(shadow);
                }
            }
        }

        protected void appendShading(XmlElement parent, ShadingDescriptor desc)
        {
            XmlElement shd = _nodeFactory.CreateElement("w", "shd", OpenXmlNamespaces.WordprocessingML);

            //fill color
            XmlAttribute fill = _nodeFactory.CreateAttribute("w", "fill", OpenXmlNamespaces.WordprocessingML);
            if (desc.cvBack != 0)
                fill.Value = new RGBColor((int)desc.cvBack, RGBColor.ByteOrder.RedLast).SixDigitHexCode;
            else
                fill.Value = desc.icoBack.ToString();
            shd.Attributes.Append(fill);

            //foreground color
            XmlAttribute color = _nodeFactory.CreateAttribute("w", "color", OpenXmlNamespaces.WordprocessingML);
            if (desc.cvFore != 0)
                color.Value = new RGBColor((int)desc.cvFore, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
            else
                color.Value = desc.icoFore.ToString();
            shd.Attributes.Append(color);

            //pattern
            XmlAttribute val = _nodeFactory.CreateAttribute("w", "val", OpenXmlNamespaces.WordprocessingML);
            val.Value = getShadingPattern(desc);
            shd.Attributes.Append(val);

            parent.AppendChild(shd);
        }

        protected string getBorderType(byte type)
        {
            switch (type)
            {
                case 0:
                    return "none";
                case 1:
                    return "single";
                case 2:
                    return "thick";
                case 3:
                    return "double";
                case 4:
                    //unused
                    return "none";
                case 5:
                    return "hairline";
                case 6:
                    return "dotted";
                case 7:
                    return "dashed";
                case 8:
                    return "dotDash";
                case 9:
                    return "dotDotDash";
                case 10:
                    return "triple";
                case 11:
                    return "thinThickSmallGap";
                case 12:
                    return "thickThinSmallGap";
                case 13:
                    return "thinThickThinSmallGap";
                case 14:
                    return "thinThickMediumGap";
                case 15:
                    return "thickThinMediumGap";
                case 16:
                    return "thinThickThinMediumGap";
                case 17:
                    return "thinThickLargeGap";
                case 18:
                    return "thickThinLargeGap";
                case 19:
                    return "thinThickThinLargeGap";
                case 20:
                    return "wave";
                case 21:
                    return "doubleWave";
                case 22:
                    return "dashSmallGap";
                case 23:
                    return "dashDotStroked";
                case 24:
                    return "threeDEmboss";
                case 25:
                    return "threeDEngrave";
                default:
                    return "none";
            }
        }

        private string getShadingPattern(ShadingDescriptor shd)
        {
            string pattern = "";
            switch (shd.ipat)
            {
                case ShadingDescriptor.ShadingPattern.Automatic:
                    pattern = "auto";
                    break;
                case ShadingDescriptor.ShadingPattern.Solid:
                    pattern = "solid";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_5:
                    pattern = "pct5";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_10:
                    pattern = "pct10";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_20:
                    pattern = "pct20";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_25:
                    pattern = "pct25";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_30:
                    pattern = "pct30";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_40:
                    pattern = "pct40";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_50:
                    pattern = "pct50";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_60:
                    pattern = "pct60";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_70:
                    pattern = "pct70";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_75:
                    pattern = "pct75";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_80:
                    pattern = "pct80";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_90:
                    pattern = "pct90";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkHorizontal:
                    pattern = "thinHorzStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkVertical:
                    pattern = "thinVertStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.DarkForwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkBackwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkCross:
                    break;
                case ShadingDescriptor.ShadingPattern.DarkDiagonalCross:
                    pattern = "thinDiagCross";
                    break;
                case ShadingDescriptor.ShadingPattern.Horizontal:
                    pattern = "horzStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.Vertical:
                    pattern = "vertStripe";
                    break;
                case ShadingDescriptor.ShadingPattern.ForwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.BackwardDiagonal:
                    break;
                case ShadingDescriptor.ShadingPattern.Cross:
                    break;
                case ShadingDescriptor.ShadingPattern.DiagonalCross:
                    pattern = "diagCross";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_2_5:
                    pattern = "pct5";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_7_5:
                    pattern = "pct10";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_12_5:
                    pattern = "pct12";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_15:
                    pattern = "pct15";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_17_5:
                    pattern = "pct15";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_22_5:
                    pattern = "pct20";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_27_5:
                    pattern = "pct30";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_32_5:
                    pattern = "pct35";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_35:
                    pattern = "pct35";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_37_5:
                    pattern = "pct37";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_42_5:
                    pattern = "pct40";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_45:
                    pattern = "pct45";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_47_5:
                    pattern = "pct45";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_52_5:
                    pattern = "pct50";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_55:
                    pattern = "pct55";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_57_5:
                    pattern = "pct55";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_62_5:
                    pattern = "pct62";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_65:
                    pattern = "pct65";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_67_5:
                    pattern = "pct65";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_72_5:
                    pattern = "pct70";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_77_5:
                    pattern = "pct75";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_82_5:
                    pattern = "pct80";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_85:
                    pattern = "pct85";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_87_5:
                    pattern = "pct87";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_92_5:
                    pattern = "pct90";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_95:
                    pattern = "pct95";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_97_5:
                    pattern = "pct95";
                    break;
                case ShadingDescriptor.ShadingPattern.Percent_97:
                    pattern = "pct95";
                    break;
                default:
                    pattern = "nil";
                    break;
            }
            return pattern;
        }

    }
}
