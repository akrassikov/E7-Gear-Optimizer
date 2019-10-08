using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tesseract;

namespace E7_Gear_Optimizer.Ocr
{
    public class OcrParser
    {
        public Item ParseItemFromClipboard()
        {
            if (Clipboard.ContainsImage())
            {
                var clipboardImage = Clipboard.GetImage();
                var bmp = new Bitmap(clipboardImage);
                return ParseImage(bmp);
            }

            return null;
        }

        public Item ParseItemFromFile(string pathToFile)
        {
            var bmp = new Bitmap(pathToFile);
            return ParseImage(bmp);
        }

        private Item ParseImage(Bitmap bmp)
        {
            // preprocess the image to increase OCR accuracy
            bmp = ImageManipulation.InvertImageColors(bmp);
            bmp = ImageManipulation.MakeGrayscale(bmp);
            bmp = ResizeTo1080(bmp);
            var region = CalculateRegion(bmp);

            // parse text from image
            var text = TesseractService.ParseText(bmp, region);
            bmp.Dispose();

            var item = ParseOcrData(text);
            return item;
        }

        private Bitmap ResizeTo1080(Bitmap bmp)
        {
            // sizes
            // 2436x1125 = 19.5 : 9 ratio
            // 2280x1080 = 19   : 9 ratio
            // 2220x1080 = 18.5 : 9 ratio
            // 2200x1080 = 18.33: 9 ratio
            // 2160x1080 = 18   : 9 ratio
            // 1920x1080 = 16   : 9 ratio

            // determine the ratio
            var ratio = bmp.Width / bmp.Height;

            var multiplier = 1080.0 / bmp.Height;

            if (bmp.Height < 1080)
            {
                var newWidth = (int)Math.Round(bmp.Width * multiplier);

                return ImageManipulation.ResizeImage(bmp, new Size(newWidth, 1080));
            }
            else if (bmp.Height > 1080)
            {
                var newWidth = (int)Math.Round(bmp.Width / multiplier);

                return ImageManipulation.ResizeImage(bmp, new Size(newWidth, 1080));
            }

            return bmp;
        }

        private Rect CalculateRegion(Bitmap bmp)
        {
            // assuming at least 16:9 ratio and 1080px height
            // aiming to crop 720x760 image

            var x = (bmp.Width - 720) / 2;
            var y = (bmp.Height - 760) / 2;

            return new Rect(x, y, 440, 760);
        }

        private Item ParseOcrData(string text)
        {
            string[] lines = text.Split(
                new[] { "\r\n", "\r", "\n" },
                StringSplitOptions.RemoveEmptyEntries
            );
            var curLine = 0;

            var item = new Item();

            var typeAndGradePattern = @"(Normal|Good|Rare|Heroic|Epic) (Weapon|Helmet|Armor|Necklace|Ring|Boots)";
            var statPattern = @"(Attack|Defense|Health|Critical Hit Chance|Critical Hit Damage|Effectiveness|Effect Resistance|Speed) (\d{1,4})(%?)";
            var setPattern = @"(Attack|Defense|Health|Effectiveness|Critical|Lifesteal|Effect Resistance|Counter|Rage|Destruction|Immunity|Unity) (Set|set)";

            // type and grade
            var typeAndGradeRgx = new Regex(typeAndGradePattern);
            while (curLine < lines.Length)
            {
                var match = typeAndGradeRgx.Match(lines[curLine]);
                if (match.Success)
                {
                    var substring = match.Value;
                    item.Type = ParseItemType(substring);
                    item.Grade = ParseGrade(substring);
                    break;
                }
                curLine++;
            }

            // enhancement (+1 -> +15)

            // stats and
            // set (attack, defense, health, effectiveness, resist, counter, rage, destruction, immunity)
            var stats = new List<Stat>();
            var statRgx = new Regex(statPattern);
            var setRgx = new Regex(setPattern);

            while (curLine < lines.Length)
            {
                var match = statRgx.Match(lines[curLine]);
                if (match.Success)
                {
                    var substring = match.Value;
                    if (stats.Count < 5)
                    {
                        stats.Add(ParseStat(substring));
                    }
                }
                else
                {
                    match = setRgx.Match(lines[curLine]);
                    if (match.Success)
                    {
                        // reached the end of the stats and found the set type
                        var substring = match.Value;
                        item.Set = ParseSet(substring);
                        item.Main = stats.First();
                        if (stats.Count > 1)
                        {
                            item.SubStats = stats.GetRange(1, stats.Count - 1).ToArray();
                        }
                        break;
                    }
                }
                curLine++;
            }

            return item;
        }

        private Grade ParseGrade(string substring)
        {
            // grade (normal, good, rare, heroic, epic)
            if (substring.Contains("Normal"))
            {
                return Grade.Normal;
            }
            else if (substring.Contains("Good"))
            {
                return Grade.Good;
            }
            else if (substring.Contains("Rare"))
            {
                return Grade.Rare;
            }
            else if (substring.Contains("Heroic"))
            {
                return Grade.Heroic;
            }
            else if (substring.Contains("Epic"))
            {
                return Grade.Epic;
            }

            return Grade.Epic;
        }

        private ItemType ParseItemType(string substring)
        {
            // type (weapon, helmet, armor, necklace, ring, boots)
            if (substring.Contains("Weapon"))
            {
                return ItemType.Weapon;
            }
            else if (substring.Contains("Helmet"))
            {
                return ItemType.Helmet;
            }
            else if (substring.Contains("Armor"))
            {
                return ItemType.Armor;
            }
            else if (substring.Contains("Necklace"))
            {
                return ItemType.Necklace;
            }
            else if (substring.Contains("Ring"))
            {
                return ItemType.Ring;
            }
            else if (substring.Contains("Boots"))
            {
                return ItemType.Boots;
            }

            return ItemType.Weapon;
        }

        private Stat ParseStat(string substring)
        {
            var split = substring.Split(' ');

            if (substring.EndsWith("%"))
            {
                // percent stat
                var valueString = split[split.Length - 1];
                var value = float.Parse(valueString.Substring(0, valueString.Length - 1));

                if (substring.Contains("Attack"))
                {
                    return new Stat(Stats.ATKPercent, value);
                }
                else if (substring.Contains("Defense"))
                {
                    return new Stat(Stats.DEFPercent, value);
                }
                else if (substring.Contains("Health"))
                {
                    return new Stat(Stats.HPPercent, value);
                }
                else if (substring.Contains("Effectiveness"))
                {
                    return new Stat(Stats.EFF, value);
                }
                else if (substring.Contains("Effect Resist"))
                {
                    return new Stat(Stats.RES, value);
                }
                else if (substring.Contains("Critical Hit Chance"))
                {
                    return new Stat(Stats.Crit, value);
                }
                else if (substring.Contains("Critical Hit Damage"))
                {
                    return new Stat(Stats.CritDmg, value);
                }
            }
            else
            {
                // flat stat
                var valueString = split[split.Length - 1];
                var value = float.Parse(valueString);

                if (substring.Contains("Attack"))
                {
                    return new Stat(Stats.ATK, value);
                }
                else if (substring.Contains("Defense"))
                {
                    return new Stat(Stats.DEF, value);
                }
                else if (substring.Contains("Health"))
                {
                    return new Stat(Stats.HP, value);
                }
                else if (substring.Contains("Speed"))
                {
                    return new Stat(Stats.SPD, value);
                }
            }

            throw new Exception("Could not parse given stat");
        }

        private Set ParseSet(string input)
        {
            var substring = input.ToLower();
            if (substring.Contains("attack set"))
            {
                return Set.Attack;
            }
            else if (substring.Contains("counter set"))
            {
                return Set.Counter;
            }
            else if (substring.Contains("critical set"))
            {
                return Set.Crit;
            }
            else if (substring.Contains("defense set"))
            {
                return Set.Def;
            }
            else if (substring.Contains("destruction set"))
            {
                return Set.Destruction;
            }
            else if (substring.Contains("health set"))
            {
                return Set.Health;
            }
            else if (substring.Contains("effectiveness set"))
            {
                return Set.Hit;
            }
            else if (substring.Contains("immunity set"))
            {
                return Set.Immunity;
            }
            else if (substring.Contains("fifesteal set"))
            {
                return Set.Lifesteal;
            }
            else if (substring.Contains("rage set"))
            {
                return Set.Rage;
            }
            else if (substring.Contains("effect resistance set"))
            {
                return Set.Resist;
            }
            else if (substring.Contains("speed set"))
            {
                return Set.Speed;
            }
            else if (substring.Contains("unity set"))
            {
                return Set.Unity;
            }

            return Set.Speed;
        }
    }
}
