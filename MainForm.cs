using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ShadowEdit
{
    public partial class MainForm : Form
    {
        FileStream reader;
        bool saveLoaded = false;
        List<RadioButton> itemSlots = new List<RadioButton>();
        List<NumericUpDown> itemValues = new List<NumericUpDown>();
        int[] runnerOffsets = { 0x0, 0x100, 0x200 };
        int runnerOffset = 0;
        int[] spellOffsets = { 0x88, 0x66, 0x44, 0x22, 0x0 };
        int spellOffset = 0;
        int runners = 0;

        string[] area0 =
        {
            "00 - Seattle Gen. Hospital",
            "01 - Icarus Descending",
            "02 - Matchsticks",
            "03 - Gates Undersound",
            "04 - Mitsuhama Computer Technologies",
            "05 - Fuchi's local office",
            "06 - Space Needle",
            "07 - Roscoe the Fixer",
            "08 - ???"
        };

        string[] area1 = 
        {
            "00 - The Wanderer",
            "01 - Tarislar City Inn",
            "02 - Ares' Corporate Offices",
            "03 - Riannon's",
            "04 - Underground 93",
            "05 - Crime Mall",
            "06 - Dr. Bob's Quickstitch",
            "07 - Tarislar Garden Apartment Complex",
            "08 - ???",
            "09 - Ork Gang"
        };

        string[] area2 = 
        {
            "00 - Hollywood Correctional Facilities",
            "01 - Little Chiba",
            "02 - Stoker's Coffin Motel",
            "03 - Shiawase Nuke Plant",
            "04 - Jump House",
            "05 - Jackal's Lantern",
            "06 - Boris' Greenhouse",
            "07 - Rat's Nest",
            "08 - ???",
            "09 - Halloweeners",
            "0A - ???",
            "0B - Ares Weapon Emporium"
        };

        string[] area3 = 
        {
            "00 - Club Penumbra",
            "01 - Wylie's Gala Inn",
            "02 - Fuchi Industrial Electronics",
            "03 - Mitsuhama's local office",
            "04 - Eye Fivers",
            "05 - Big Rhino",
            "06 - Lone Star Security"
        };

        string[] area4 = 
        {
            "00 - Council Island Inn",
            "01 - Friendship Restaurant",
            "02 - Passport Lodge",
            "03 - Medicine Lodge Hollow",
            "04 - Council Island Hospital",
            "05 - Ork Embassy"
        };

        string[] area5 = 
        {
            "00 - ???",
            "01 - Frag Grenade",
            "02 - Wire-Masters",
            "03 - Merlin's Lore",
            "04 - Microtronics",
            "05 - Weapons World",
            "06 - Renraku Offices"
        };

        string[] area6 = 
        {
            "00 - ???"
        };

        // TODO: stop being lazy and rename form controls :P
        public MainForm()
        {
            InitializeComponent();

            // Load radio buttons into list for iteration for accessories editor
            itemSlots.Add(RadioButtonSlot1);
            itemSlots.Add(RadioButtonSlot2);
            itemSlots.Add(RadioButtonSlot3);
            itemSlots.Add(RadioButtonSlot4);
            itemSlots.Add(RadioButtonSlot5);
            itemSlots.Add(RadioButtonSlot6);
            itemSlots.Add(RadioButtonSlot7);
            itemSlots.Add(RadioButtonSlot8);

            // Load NumericUpDowns into list for iteration for accessories editor
            itemValues.Add(NumericUpDownInventory1);
            itemValues.Add(NumericUpDownInventory2);
            itemValues.Add(NumericUpDownInventory3);
            itemValues.Add(NumericUpDownInventory4);
            itemValues.Add(NumericUpDownInventory5);
            itemValues.Add(NumericUpDownInventory6);
            itemValues.Add(NumericUpDownInventory7);
            itemValues.Add(NumericUpDownInventory8);
        }

        private void ButtoBrowse_Click(object sender, EventArgs e)
        {
            FileDialog dlg = new OpenFileDialog();
            dlg.Title = "Open Save State File...";
            dlg.ShowDialog();

            TextBoxFile.Text = dlg.FileName;
        }

        private void LoadBasics()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Stance
            reader.Seek(0x25C5 + runnerOffset, SeekOrigin.Begin);
            TrackBarPosture.Value = reader.ReadByte();

            // Race
            reader.Seek(0x2612 + runnerOffset, SeekOrigin.Begin);
            ComboBoxRace.SelectedIndex = reader.ReadByte();

            // Archetype
            reader.Seek(0x2613 + runnerOffset, SeekOrigin.Begin);
            ComboBoxArchetype.SelectedIndex = reader.ReadByte();

            // Ammo Left In Gun
            reader.Seek(0x25C6 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownAmmo.Value = reader.ReadByte();

            // Clips
            reader.Seek(0x25C7 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownClips.Value = reader.ReadByte();

            // Current Weapon
            bool isSpell;
            reader.Seek(0x25CE + runnerOffset, SeekOrigin.Begin);
            if (reader.ReadByte() == 255)
                isSpell = true;
            else
                isSpell = false;
            int currentWeapon;
            reader.Seek(0x25CF + runnerOffset, SeekOrigin.Begin);
            currentWeapon = reader.ReadByte();
            if (currentWeapon == 0x15)
                ComboBoxCurrentWeapon.SelectedIndex = 16;
            if (isSpell)
                ComboBoxCurrentWeapon.SelectedIndex = currentWeapon + 16;
            else
                ComboBoxCurrentWeapon.SelectedIndex = currentWeapon;

            // Current Armor
            reader.Seek(0x25EE + runnerOffset, SeekOrigin.Begin);
            ComboBoxCurrentArmor.SelectedIndex = reader.ReadByte();

            // Nuyen
            reader.Seek(0x12076, SeekOrigin.Begin);
            int nuyen = 0;
            float result = 0;
            for (int i = 0; i < 4; i++)
            {
                nuyen = reader.ReadByte();
                result += nuyen << (3 - i) * 8;
            }
            NumericUpDownNuyen.Value = Convert.ToInt32(result);

            // Karma
            reader.Seek(0x2605 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownKarma.Text = reader.ReadByte().ToString();

            reader.Close();
        }

        private void LoadInventory()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Inventory
            reader.Seek(0x25DE + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory1.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25DF + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory2.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E0 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory3.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E1 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory4.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E2 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory5.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E3 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory6.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E4 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory7.SelectedIndex = reader.ReadByte();
            reader.Seek(0x25E5 + runnerOffset, SeekOrigin.Begin);
            ComboBoxInventory8.SelectedIndex = reader.ReadByte();

            // Inventory Mods / Quantity
            reader.Seek(0x25E6 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory1.Text = reader.ReadByte().ToString();
            reader.Seek(0x25E7 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory2.Text = reader.ReadByte().ToString();
            reader.Seek(0x25E8 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory3.Text = reader.ReadByte().ToString();
            reader.Seek(0x25E9 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory4.Text = reader.ReadByte().ToString();
            reader.Seek(0x25EA + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory5.Text = reader.ReadByte().ToString();
            reader.Seek(0x25EB + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory6.Text = reader.ReadByte().ToString();
            reader.Seek(0x25EC + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory7.Text = reader.ReadByte().ToString();
            reader.Seek(0x25ED + runnerOffset, SeekOrigin.Begin);
            NumericUpDownInventory8.Text = reader.ReadByte().ToString();

            // Load up the item images in the quasi-inventory section
            ApplyItemImages();

            reader.Close();
        }

        private void LoadAttributes()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Attributes
            reader.Seek(0x25EF + runnerOffset, SeekOrigin.Begin);
            NumericUpDownBody.Value = reader.ReadByte();
            reader.Seek(0x25F0 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownQuickness.Value = reader.ReadByte();
            reader.Seek(0x25F1 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownStrength.Value = reader.ReadByte();
            reader.Seek(0x25F2 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownCharisma.Value = reader.ReadByte();
            reader.Seek(0x25F3 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownIntelligence.Value = reader.ReadByte();
            reader.Seek(0x25F4 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownWillpower.Value = reader.ReadByte();
            reader.Seek(0x25F6 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownEssence1.Value = reader.ReadByte();
            reader.Seek(0x2606 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownEssence2.Value = reader.ReadByte();
            reader.Seek(0x25F7 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownMagic.Value = reader.ReadByte();

            reader.Close();
        }

        private void LoadSkills()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Skills
            reader.Seek(0x25F8 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownSorcery.Value = reader.ReadByte();
            reader.Seek(0x25F9 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownFirearms.Value = reader.ReadByte();
            reader.Seek(0x25FA + runnerOffset, SeekOrigin.Begin);
            NumericUpDownPistols.Value = reader.ReadByte();
            reader.Seek(0x25FB + runnerOffset, SeekOrigin.Begin);
            NumericUpDownSMGs.Value = reader.ReadByte();
            reader.Seek(0x25FC + runnerOffset, SeekOrigin.Begin);
            NumericUpDownShotguns.Value = reader.ReadByte();
            reader.Seek(0x25FD + runnerOffset, SeekOrigin.Begin);
            NumericUpDownMelee.Value = reader.ReadByte();
            reader.Seek(0x25FF + runnerOffset, SeekOrigin.Begin);
            NumericUpDownThrowing.Value = reader.ReadByte();
            reader.Seek(0x2600 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownComputer.Value = reader.ReadByte();
            reader.Seek(0x2601 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownBioTech.Value = reader.ReadByte();
            reader.Seek(0x2602 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownElectronics.Value = reader.ReadByte();
            reader.Seek(0x2603 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownReputation.Value = reader.ReadByte();
            reader.Seek(0x2604 + runnerOffset, SeekOrigin.Begin);
            NumericUpDownNegotiation.Value = reader.ReadByte();

            reader.Close();
        }

        private void LoadCyberdeck()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Cyberdeck presence flag
            reader.Seek(0x2607, SeekOrigin.Begin);
            bool present = Convert.ToBoolean(reader.ReadByte());
            CheckBoxHasCyberdeck.Checked = present;

            // Cyberdeck programs
            reader.Seek(0x12044, SeekOrigin.Begin);
            NumericUpDownAttack.Value = reader.ReadByte();
            reader.Seek(0x12045, SeekOrigin.Begin);
            NumericUpDownSlow.Value = reader.ReadByte();
            reader.Seek(0x12046, SeekOrigin.Begin);
            NumericUpDownDegrade.Value = reader.ReadByte();
            reader.Seek(0x12047, SeekOrigin.Begin);
            NumericUpDownRebound.Value = reader.ReadByte();
            reader.Seek(0x12048, SeekOrigin.Begin);
            NumericUpDownMedic.Value = reader.ReadByte();
            reader.Seek(0x12049, SeekOrigin.Begin);
            NumericUpDownShield.Value = reader.ReadByte();
            reader.Seek(0x1204A, SeekOrigin.Begin);
            NumericUpDownSmoke.Value = reader.ReadByte();
            reader.Seek(0x1204B, SeekOrigin.Begin);
            NumericUpDownMirrors.Value = reader.ReadByte();
            reader.Seek(0x1204C, SeekOrigin.Begin);
            NumericUpDownSleaze.Value = reader.ReadByte();
            reader.Seek(0x1204D, SeekOrigin.Begin);
            NumericUpDownDeception.Value = reader.ReadByte();
            reader.Seek(0x1204E, SeekOrigin.Begin);
            NumericUpDownRelocation.Value = reader.ReadByte();
            reader.Seek(0x1204F, SeekOrigin.Begin);
            NumericUpDownAnalyze.Value = reader.ReadByte();

            // Cyberdeck Stats
            reader.Seek(0x12036, SeekOrigin.Begin);
            NumericUpDownMPCPRating.Value = reader.ReadByte();
            reader.Seek(0x12037, SeekOrigin.Begin);
            NumericUpDownHardeningRating.Value = reader.ReadByte();
            reader.Seek(0x1203C, SeekOrigin.Begin);
            NumericUpDownLoadIOSpeed.Value = reader.ReadByte();
            reader.Seek(0x1203D, SeekOrigin.Begin);
            NumericUpDownResponse.Value = reader.ReadByte();
            reader.Seek(0x1203E, SeekOrigin.Begin);
            NumericUpDownBodyRating.Value = reader.ReadByte();
            reader.Seek(0x1203F, SeekOrigin.Begin);
            NumericUpDownEvasionRating.Value = reader.ReadByte();
            reader.Seek(0x12040, SeekOrigin.Begin);
            NumericUpDownMaskingRating.Value = reader.ReadByte();
            reader.Seek(0x12041, SeekOrigin.Begin);
            NumericUpDownSensorRating.Value = reader.ReadByte();

            // Cyberdeck memory
            reader.Seek(0x12038, SeekOrigin.Begin);
            int memory = 0;
            int result = 0;
            result = 0;
            for (int i = 0; i < 2; i++)
            {
                memory = reader.ReadByte();
                result += memory << (1 - i) * 8;
            }
            NumericUpDownMemory.Value = Convert.ToInt32(result);
            // Cyberdeck storage
            reader.Seek(0x1203A, SeekOrigin.Begin);
            int storage = 0;
            result = 0;
            for (int i = 0; i < 2; i++)
            {
                storage = reader.ReadByte();
                result += storage << (1 - i) * 8;
            }
            NumericUpDownStorage.Value = Convert.ToInt32(result);

            // Cyberdeck brand
            reader.Seek(0x12043, SeekOrigin.Begin);
            int brand = reader.ReadByte();
            if (brand >= 5)
                ComboBoxCyberdeckBrand.SelectedIndex = 5;
            else if (brand <= 4)
                ComboBoxCyberdeckBrand.SelectedIndex = brand;

            reader.Close();
        }

        private void LoadSpellbook()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Spellbook
            int spellCheck;
            reader.Seek(0x11EFC - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell1.SelectedIndex = 14;
            else
                ComboBoxSpell1.SelectedIndex = spellCheck;
            reader.Seek(0x11EFD - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel1.Value = reader.ReadByte();
            reader.Seek(0x11EFE - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell2.SelectedIndex = 14;
            else
                ComboBoxSpell2.SelectedIndex = spellCheck;
            reader.Seek(0x11EFF - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel2.Value = reader.ReadByte();
            reader.Seek(0x11F00 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell3.SelectedIndex = 14;
            else
                ComboBoxSpell3.SelectedIndex = spellCheck;
            reader.Seek(0x11F01 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel3.Value = reader.ReadByte();
            reader.Seek(0x11F02 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell4.SelectedIndex = 14;
            else
                ComboBoxSpell4.SelectedIndex = spellCheck;
            reader.Seek(0x11F03 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel4.Value = reader.ReadByte();
            reader.Seek(0x11F04 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell5.SelectedIndex = 14;
            else
                ComboBoxSpell5.SelectedIndex = spellCheck;
            reader.Seek(0x11F05 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel5.Value = reader.ReadByte();
            reader.Seek(0x11F06 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell6.SelectedIndex = 14;
            else
                ComboBoxSpell6.SelectedIndex = spellCheck;
            reader.Seek(0x11F07 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel6.Value = reader.ReadByte();
            reader.Seek(0x11F08 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell7.SelectedIndex = 14;
            else
                ComboBoxSpell7.SelectedIndex = spellCheck;
            reader.Seek(0x11F09 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel7.Value = reader.ReadByte();
            reader.Seek(0x11F0A - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell8.SelectedIndex = 14;
            else
                ComboBoxSpell8.SelectedIndex = spellCheck;
            reader.Seek(0x11F0B - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel8.Value = reader.ReadByte();
            reader.Seek(0x11F0C - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell9.SelectedIndex = 14;
            else
                ComboBoxSpell9.SelectedIndex = spellCheck;
            reader.Seek(0x11F0D - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel9.Value = reader.ReadByte();
            reader.Seek(0x11F0E - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell10.SelectedIndex = 14;
            else
                ComboBoxSpell10.SelectedIndex = spellCheck;
            reader.Seek(0x11F0F - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel10.Value = reader.ReadByte();
            reader.Seek(0x11F10 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell11.SelectedIndex = 14;
            else
                ComboBoxSpell11.SelectedIndex = spellCheck;
            reader.Seek(0x11F11 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel11.Value = reader.ReadByte();
            reader.Seek(0x11F12 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell12.SelectedIndex = 14;
            else
                ComboBoxSpell12.SelectedIndex = spellCheck;
            reader.Seek(0x11F13 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel12.Value = reader.ReadByte();
            reader.Seek(0x11F14 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell13.SelectedIndex = 14;
            else
                ComboBoxSpell13.SelectedIndex = spellCheck;
            reader.Seek(0x11F15 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel13.Value = reader.ReadByte();
            reader.Seek(0x11F16 - spellOffset, SeekOrigin.Begin);
            spellCheck = reader.ReadByte();
            if (spellCheck == 255)
                ComboBoxSpell14.SelectedIndex = 14;
            else
                ComboBoxSpell14.SelectedIndex = spellCheck;
            reader.Seek(0x11F17 - spellOffset, SeekOrigin.Begin);
            TrackBarSpellLevel14.Value = reader.ReadByte();

            // Spellbook total spells
            reader.Seek(0x11EFB - spellOffset, SeekOrigin.Begin);
            NumericUpDownSpells.Value = reader.ReadByte();

            reader.Close();
        }

        private void LoadCurrentRun()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int johnson;
            int runType;
            bool ghoulBounty = false;
            bool matrixRun = false;

            // Johnson
            reader.Seek(0x12064, SeekOrigin.Begin);
            johnson = reader.ReadByte();
            if (johnson == 0xFF)
                ComboBoxCurrentJohnson.SelectedIndex = 5;
            else
                ComboBoxCurrentJohnson.SelectedIndex = johnson;

            // Run Type
            reader.Seek(0x12065, SeekOrigin.Begin);
            runType = reader.ReadByte();
            if (runType == 0)
                ghoulBounty = true;
            if (runType == 6)
                matrixRun = true;
            if (runType == 0xFF)
                ComboBoxRunType.SelectedIndex = 7;
            else
                ComboBoxRunType.SelectedIndex = runType;

            if (johnson == 0xFF || runType == 0xFF)
            {
                // Don't load anything
                NumericUpDownPayment.Value = 0;
            }
            else
            {
                // Area assignment
                if (!matrixRun)
                {
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    ComboBoxRunSourceArea.SelectedIndex = reader.ReadByte();
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    ComboBoxRunSourceBuilding.SelectedIndex = reader.ReadByte();
                    if (!ghoulBounty)
                    {
                        reader.Seek(0x12069, SeekOrigin.Begin);
                        ComboBoxRunDestinationArea.SelectedIndex = reader.ReadByte();
                        reader.Seek(0x1206A, SeekOrigin.Begin);
                        ComboBoxRunDestinationBuilding.SelectedIndex = reader.ReadByte();
                    }
                }

                // Matrix run type/system/listed
                reader.Seek(0x12066, SeekOrigin.Begin);
                ComboBoxMatrixRunType.SelectedIndex = reader.ReadByte();
                reader.Seek(0x12067, SeekOrigin.Begin);
                ComboBoxMatrixRunListedSystem.SelectedIndex = reader.ReadByte();
                reader.Seek(0x12068, SeekOrigin.Begin);
                ComboBoxMatrixRunSystem.SelectedIndex = reader.ReadByte();

                // Escort/Extract client name
                reader.Seek(0x12068, SeekOrigin.Begin);
                ComboBoxClientName.SelectedIndex = reader.ReadByte();

                // Payment
                reader.Seek(0x1206C, SeekOrigin.Begin);
                int payment = 0;
                float result = 0;
                for (int i = 0; i < 2; i++)
                {
                    payment = reader.ReadByte();
                    result += payment << (1 - i) * 8;
                }
                NumericUpDownPayment.Value = Convert.ToInt32(result);

                // Flags
                reader.Seek(0x12094, SeekOrigin.Begin);
                NumericUpDownRunFlags.Value = reader.ReadByte();
            }

            reader.Close();
        }

        private void LoadCyberware()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1, value2;

            // Uncheck everything first to prevent conflicts
            CheckBoxMuscleReplacement4.Checked = false;
            CheckBoxDermalPlating1.Checked = false;
            CheckBoxDermalPlating2.Checked = false;
            CheckBoxDermalPlating3.Checked = false;
            CheckBoxWiredReflexes1.Checked = false;
            CheckBoxWiredReflexes2.Checked = false;
            CheckBoxWiredReflexes3.Checked = false;
            CheckBoxDatajack.Checked = false;
            CheckBoxCyberEyes.Checked = false;
            CheckBoxHandRazors.Checked = false;
            CheckBoxSpurs.Checked = false;
            CheckBoxSmartlink.Checked = false;
            CheckBoxMuscleReplacement1.Checked = false;
            CheckBoxMuscleReplacement2.Checked = false;
            CheckBoxMuscleReplacement3.Checked = false;

            // First set
            reader.Seek(0x2608 + runnerOffset, SeekOrigin.Begin);
            value1 = reader.ReadByte();
            if ((value1 & 1) == 1)
                CheckBoxMuscleReplacement4.Checked = true;
            if ((value1 & 2) == 2)
                CheckBoxDermalPlating1.Checked = true;
            if ((value1 & 4) == 4)
                CheckBoxDermalPlating2.Checked = true;
            if ((value1 & 8) == 8)
                CheckBoxDermalPlating3.Checked = true;
            if ((value1 & 16) == 16)
                CheckBoxWiredReflexes1.Checked = true;
            if ((value1 & 32) == 32)
                CheckBoxWiredReflexes2.Checked = true;
            if ((value1 & 64) == 64)
                CheckBoxWiredReflexes3.Checked = true;

            // Second set
            reader.Seek(0x2609 + runnerOffset, SeekOrigin.Begin);
            value2 = reader.ReadByte();
            if ((value2 & 1) == 1)
                CheckBoxDatajack.Checked = true;
            if ((value2 & 2) == 2)
                CheckBoxCyberEyes.Checked = true;
            if ((value2 & 4) == 4)
                CheckBoxHandRazors.Checked = true;
            if ((value2 & 8) == 8)
                CheckBoxSpurs.Checked = true;
            if ((value2 & 16) == 16)
                CheckBoxSmartlink.Checked = true;
            if ((value2 & 32) == 32)
                CheckBoxMuscleReplacement1.Checked = true;
            if ((value2 & 64) == 64)
                CheckBoxMuscleReplacement2.Checked = true;
            if ((value2 & 128) == 128)
                CheckBoxMuscleReplacement3.Checked = true;

            reader.Close();
        }

        private void LoadGroupItems()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1, value2;

            // Uncheck everything first to prevent conflicts
            checkBox15.Checked = false;
            checkBox14.Checked = false;
            checkBox13.Checked = false;
            checkBox12.Checked = false;
            checkBox11.Checked = false;
            checkBox10.Checked = false;
            checkBox9.Checked = false;
            checkBox8.Checked = false;
            checkBox7.Checked = false;
            checkBox6.Checked = false;
            checkBox5.Checked = false;
            checkBox4.Checked = false;
            checkBox3.Checked = false;
            checkBox2.Checked = false;
            checkBox1.Checked = false;

            // First set
            reader.Seek(0x1209C, SeekOrigin.Begin);
            value1 = reader.ReadByte();
            if ((value1 & 1) == 1)
                checkBox15.Checked = true;
            if ((value1 & 2) == 2)
                checkBox14.Checked = true;
            if ((value1 & 4) == 4)
                checkBox13.Checked = true;
            if ((value1 & 8) == 8)
                checkBox12.Checked = true;
            if ((value1 & 16) == 16)
                checkBox11.Checked = true;
            if ((value1 & 32) == 32)
                checkBox10.Checked = true;
            if ((value1 & 64) == 64)
                checkBox9.Checked = true;

            // Second set
            reader.Seek(0x1209D, SeekOrigin.Begin);
            value2 = reader.ReadByte();
            if ((value2 & 1) == 1)
                checkBox8.Checked = true;
            if ((value2 & 2) == 2)
                checkBox7.Checked = true;
            if ((value2 & 4) == 4)
                checkBox6.Checked = true;
            if ((value2 & 8) == 8)
                checkBox5.Checked = true;
            if ((value2 & 16) == 16)
                checkBox4.Checked = true;
            if ((value2 & 32) == 32)
                checkBox3.Checked = true;
            if ((value2 & 64) == 64)
                checkBox2.Checked = true;
            if ((value2 & 128) == 128)
                checkBox1.Checked = true;

            reader.Close();
        }

        private void LoadMatrixPasscodes()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1, value2;

            // Uncheck everything first to prevent conflicts
            checkBox19.Checked = false;
            checkBox18.Checked = false;
            checkBox17.Checked = false;
            checkBox16.Checked = false;
            checkBox31.Checked = false;
            checkBox32.Checked = false;
            checkBox34.Checked = false;
            checkBox35.Checked = false;
            checkBox33.Checked = false;
            checkBox29.Checked = false;
            checkBox30.Checked = false;
            checkBox28.Checked = false;

            // First set
            reader.Seek(0x120A0, SeekOrigin.Begin);
            value1 = reader.ReadByte();
            if ((value1 & 1) == 1)
                checkBox19.Checked = true;
            if ((value1 & 2) == 2)
                checkBox18.Checked = true;
            if ((value1 & 4) == 4)
                checkBox17.Checked = true;
            if ((value1 & 8) == 8)
                checkBox16.Checked = true;

            // Second set
            reader.Seek(0x120A1, SeekOrigin.Begin);
            value2 = reader.ReadByte();
            if ((value2 & 1) == 1)
                checkBox31.Checked = true;
            if ((value2 & 2) == 2)
                checkBox32.Checked = true;
            if ((value2 & 4) == 4)
                checkBox34.Checked = true;
            if ((value2 & 8) == 8)
                checkBox35.Checked = true;
            if ((value2 & 16) == 16)
                checkBox33.Checked = true;
            if ((value2 & 32) == 32)
                checkBox29.Checked = true;
            if ((value2 & 64) == 64)
                checkBox30.Checked = true;
            if ((value2 & 128) == 128)
                checkBox28.Checked = true;

            reader.Close();
        }

        private void SaveBasics()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Stance
            reader.Seek(0x25C5 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarPosture.Value));

            // Race
            reader.Seek(0x2612 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxRace.SelectedIndex));

            // Archetype
            reader.Seek(0x2613 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxArchetype.SelectedIndex));

            // Ammo Left In Gun
            reader.Seek(0x25C6 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownAmmo.Value));

            // Clips
            reader.Seek(0x25C7 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownClips.Value));

            // Current Weapon
            bool isSpell;
            reader.Seek(0x25CE + runnerOffset, SeekOrigin.Begin);
            if (ComboBoxCurrentWeapon.SelectedIndex > 17)
            {
                reader.WriteByte(Convert.ToByte(0xFF));
                isSpell = true;
            }
            else
            {
                reader.WriteByte(Convert.ToByte(0));
                isSpell = false;
            }
            reader.Seek(0x25CF + runnerOffset, SeekOrigin.Begin);
            if (ComboBoxCurrentWeapon.SelectedIndex == 16)
                reader.WriteByte(0x15);
            if (isSpell)
                reader.WriteByte(Convert.ToByte(ComboBoxCurrentWeapon.SelectedIndex - 16));
            else
                reader.WriteByte(Convert.ToByte(ComboBoxCurrentWeapon.SelectedIndex));

            // Current Armor
            reader.Seek(0x25EE + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxCurrentArmor.SelectedIndex));

            // Nuyen
            reader.Seek(0x12076, SeekOrigin.Begin);
            int nuyen = Convert.ToInt32(NumericUpDownNuyen.Value);
            for (int i = 0; i < 4; i++)
                reader.WriteByte((byte)(nuyen >> (3 - i) * 8));

            // Karma
            reader.Seek(0x2605 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownKarma.Value));

            reader.Close();
        }

        private void SaveInventory()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Inventory
            reader.Seek(0x25DE + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory1.SelectedIndex));
            reader.Seek(0x25DF + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory2.SelectedIndex));
            reader.Seek(0x25E0 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory3.SelectedIndex));
            reader.Seek(0x25E1 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory4.SelectedIndex));
            reader.Seek(0x25E2 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory5.SelectedIndex));
            reader.Seek(0x25E3 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory6.SelectedIndex));
            reader.Seek(0x25E4 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory7.SelectedIndex));
            reader.Seek(0x25E5 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(ComboBoxInventory8.SelectedIndex));

            // Inventory Mods / Quantity
            reader.Seek(0x25E6 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory1.Value));
            reader.Seek(0x25E7 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory2.Value));
            reader.Seek(0x25E8 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory3.Value));
            reader.Seek(0x25E9 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory4.Value));
            reader.Seek(0x25EA + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory5.Value));
            reader.Seek(0x25EB + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory6.Value));
            reader.Seek(0x25EC + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory7.Value));
            reader.Seek(0x25ED + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownInventory8.Value));

            reader.Close();
        }

        private void SaveAttributes()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Attributes
            reader.Seek(0x25EF + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownBody.Value));
            reader.Seek(0x25F0 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownQuickness.Value));
            reader.Seek(0x25F1 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownStrength.Value));
            reader.Seek(0x25F2 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownCharisma.Value));
            reader.Seek(0x25F3 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownIntelligence.Value));
            reader.Seek(0x25F4 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownWillpower.Value));
            reader.Seek(0x25F6 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownEssence1.Value));
            reader.Seek(0x2606 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownEssence2.Value));
            reader.Seek(0x25F7 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMagic.Value));

            reader.Close();
        }

        private void SaveSkills()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Skills
            reader.Seek(0x25F8 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSorcery.Value));
            reader.Seek(0x25F9 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownFirearms.Value));
            reader.Seek(0x25FA + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownPistols.Value));
            reader.Seek(0x25FB + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSMGs.Value));
            reader.Seek(0x25FC + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownShotguns.Value));
            reader.Seek(0x25FD + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMelee.Value));
            reader.Seek(0x25FF + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownThrowing.Value));
            reader.Seek(0x2600 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownComputer.Value));
            reader.Seek(0x2601 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownBioTech.Value));
            reader.Seek(0x2602 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownElectronics.Value));
            reader.Seek(0x2603 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownReputation.Value));
            reader.Seek(0x2604 + runnerOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownNegotiation.Value));

            reader.Close();
        }

        private void SaveCyberdeck()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Cyberdeck presence flag
            reader.Seek(0x2607, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(CheckBoxHasCyberdeck.Checked));

            // Cyberdeck programs
            reader.Seek(0x12044, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownAttack.Value));
            reader.Seek(0x12045, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSlow.Value));
            reader.Seek(0x12046, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownDegrade.Value));
            reader.Seek(0x12047, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownRebound.Value));
            reader.Seek(0x12048, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMedic.Value));
            reader.Seek(0x12049, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownShield.Value));
            reader.Seek(0x1204A, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSmoke.Value));
            reader.Seek(0x1204B, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMirrors.Value));
            reader.Seek(0x1204C, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSleaze.Value));
            reader.Seek(0x1204D, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownDeception.Value));
            reader.Seek(0x1204E, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownRelocation.Value));
            reader.Seek(0x1204F, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownAnalyze.Value));

            // Cyberdeck Stats
            reader.Seek(0x12036, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMPCPRating.Value));
            reader.Seek(0x12037, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownHardeningRating.Value));
            // TODO: Fix IO Speed & Reponse Vars
            reader.Seek(0x1203C, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownLoadIOSpeed.Value));
            reader.Seek(0x1203D, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownResponse.Value));
            reader.Seek(0x1203E, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownBodyRating.Value));
            reader.Seek(0x1203F, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownEvasionRating.Value));
            reader.Seek(0x12040, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownMaskingRating.Value));
            reader.Seek(0x12041, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownSensorRating.Value));

            // Cyberdeck memory
            reader.Seek(0x12038, SeekOrigin.Begin);
            int memory = Convert.ToInt32(NumericUpDownMemory.Value);
            for (int i = 0; i < 2; i++)
                reader.WriteByte((byte)(memory >> (1 - i) * 8));
            // Cyberdeck storage
            reader.Seek(0x1203A, SeekOrigin.Begin);
            int storage = Convert.ToInt32(NumericUpDownStorage.Value);
            for (int i = 0; i < 2; i++)
                reader.WriteByte((byte)(storage >> (1 - i) * 8));

            // Cyberdeck brand
            reader.Seek(0x12043, SeekOrigin.Begin);
            if (ComboBoxCyberdeckBrand.SelectedIndex == 5)
                reader.WriteByte(Convert.ToByte(255));
            else if (ComboBoxCyberdeckBrand.SelectedIndex <= 4)
                reader.WriteByte(Convert.ToByte(ComboBoxCyberdeckBrand.SelectedIndex));

            reader.Close();
        }

        private void SaveSpellbook()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            // Spellbook
            reader.Seek(0x11EFC - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell1.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell1.SelectedIndex));
            reader.Seek(0x11EFD - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel1.Value));
            reader.Seek(0x11EFE - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell2.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell2.SelectedIndex));
            reader.Seek(0x11EFF - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel2.Value));
            reader.Seek(0x11F00 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell3.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell3.SelectedIndex));
            reader.Seek(0x11F01 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel3.Value));
            reader.Seek(0x11F02 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell4.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell4.SelectedIndex));
            reader.Seek(0x11F03 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel4.Value));
            reader.Seek(0x11F04 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell5.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell5.SelectedIndex));
            reader.Seek(0x11F05 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel5.Value));
            reader.Seek(0x11F06 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell6.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell6.SelectedIndex));
            reader.Seek(0x11F07 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel6.Value));
            reader.Seek(0x11F08 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell7.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell7.SelectedIndex));
            reader.Seek(0x11F09 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel7.Value));
            reader.Seek(0x11F0A - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell8.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell8.SelectedIndex));
            reader.Seek(0x11F0B - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel8.Value));
            reader.Seek(0x11F0C - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell9.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell9.SelectedIndex));
            reader.Seek(0x11F0D - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel9.Value));
            reader.Seek(0x11F0E - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell10.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell10.SelectedIndex));
            reader.Seek(0x11F0F - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel10.Value));
            reader.Seek(0x11F10 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell11.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell11.SelectedIndex));
            reader.Seek(0x11F11 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel11.Value));
            reader.Seek(0x11F12 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell12.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell12.SelectedIndex));
            reader.Seek(0x11F13 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel12.Value));
            reader.Seek(0x11F14 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell13.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell13.SelectedIndex));
            reader.Seek(0x11F15 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel13.Value));
            reader.Seek(0x11F16 - spellOffset, SeekOrigin.Begin);
            if (ComboBoxSpell14.SelectedIndex == 14)
                reader.WriteByte(0xFF);
            else
                reader.WriteByte(Convert.ToByte(ComboBoxSpell14.SelectedIndex));
            reader.Seek(0x11F17 - spellOffset, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(TrackBarSpellLevel14.Value));

            // Spellbook total spells
            int totalSpells = -1;
            List<ComboBox> spells = new List<ComboBox>();
            spells.Add(ComboBoxSpell1);
            spells.Add(ComboBoxSpell2);
            spells.Add(ComboBoxSpell3);
            spells.Add(ComboBoxSpell4);
            spells.Add(ComboBoxSpell5);
            spells.Add(ComboBoxSpell6);
            spells.Add(ComboBoxSpell7);
            spells.Add(ComboBoxSpell8);
            spells.Add(ComboBoxSpell9);
            spells.Add(ComboBoxSpell10);
            spells.Add(ComboBoxSpell11);
            spells.Add(ComboBoxSpell12);
            spells.Add(ComboBoxSpell13);
            spells.Add(ComboBoxSpell14);
            reader.Seek(0x11EFB - spellOffset, SeekOrigin.Begin);
            for (int i = 0; i <= 13; i++)
            {
                if (spells[i].SelectedIndex == 14)
                    totalSpells = totalSpells;
                else
                    totalSpells += 1;
            }
            if (CheckBoxModifySpells.Checked == true)
                NumericUpDownSpells.Value = NumericUpDownSpells.Value;
            else
                NumericUpDownSpells.Value = totalSpells;
            reader.WriteByte(Convert.ToByte(totalSpells));

            reader.Close();
        }

        private void SaveCurrentRun()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int johnson;
            int runType;

            // Johnson
            reader.Seek(0x12064, SeekOrigin.Begin);
            johnson = ComboBoxCurrentJohnson.SelectedIndex;
            if (johnson == 5)
                reader.WriteByte(Convert.ToByte(0xFF));
            else
                reader.WriteByte(Convert.ToByte(johnson));

            // Run type
            reader.Seek(0x12065, SeekOrigin.Begin);
            runType = ComboBoxRunType.SelectedIndex;
            if (runType == 7)
                reader.WriteByte(Convert.ToByte(0xFF));
            else
                reader.WriteByte(Convert.ToByte(runType));

            // Insert huge area check and such here
            switch (runType)
            {
                case 0: // Ghoul Bounty
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    break;
                case 1: // Bodyguard
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceBuilding.SelectedIndex));
                    reader.Seek(0x12069, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunDestinationArea.SelectedIndex));
                    reader.Seek(0x1206A, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunDestinationBuilding.SelectedIndex));
                    reader.Seek(0x12068, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxClientName.SelectedIndex));
                    break;
                case 2: // Courier
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceBuilding.SelectedIndex));
                    reader.Seek(0x12069, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunDestinationArea.SelectedIndex));
                    reader.Seek(0x1206A, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunDestinationBuilding.SelectedIndex));
                    break;
                case 3: // Enforcement
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceBuilding.SelectedIndex));
                    break;
                case 4: // Acquisition
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceBuilding.SelectedIndex));
                    break;
                case 5: // Extraction
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceArea.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxRunSourceBuilding.SelectedIndex));
                    reader.Seek(0x12068, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxClientName.SelectedIndex));
                    break;
                case 6: // Matrix Run
                    reader.Seek(0x12066, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxMatrixRunType.SelectedIndex));
                    reader.Seek(0x12067, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxMatrixRunListedSystem.SelectedIndex));
                    reader.Seek(0x12068, SeekOrigin.Begin);
                    reader.WriteByte(Convert.ToByte(ComboBoxMatrixRunSystem.SelectedIndex));
                    break;
            }

            // Payment
            reader.Seek(0x1206C, SeekOrigin.Begin);
            int payment = Convert.ToInt32(NumericUpDownPayment.Value);
            for (int i = 0; i < 4; i++)
                reader.WriteByte((byte)(payment >> (1 - i) * 8));

            // Flags
            reader.Seek(0x12094, SeekOrigin.Begin);
            reader.WriteByte(Convert.ToByte(NumericUpDownRunFlags.Value));

            reader.Close();
        }

        private void SaveCyberware()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1 = 0, value2 = 0;

            // First set
            reader.Seek(0x2608 + runnerOffset, SeekOrigin.Begin);
            if (CheckBoxMuscleReplacement4.Checked)
                value1 |= 1;
            if (CheckBoxDermalPlating1.Checked)
                value1 |= 2;
            if (CheckBoxDermalPlating2.Checked)
                value1 |= 4;
            if (CheckBoxDermalPlating3.Checked)
                value1 |= 8;
            if (CheckBoxWiredReflexes1.Checked)
                value1 |= 16;
            if (CheckBoxWiredReflexes2.Checked)
                value1 |= 32;
            if (CheckBoxWiredReflexes3.Checked)
                value1 |= 64;
            reader.WriteByte(Convert.ToByte(value1));

            // Second set
            reader.Seek(0x2609 + runnerOffset, SeekOrigin.Begin);
            if (CheckBoxDatajack.Checked)
                value2 |= 1;
            if (CheckBoxCyberEyes.Checked)
                value2 |= 2;
            if (CheckBoxHandRazors.Checked)
                value2 |= 4;
            if (CheckBoxSpurs.Checked)
                value2 |= 8;
            if (CheckBoxSmartlink.Checked)
                value2 |= 16;
            if (CheckBoxMuscleReplacement1.Checked)
                value2 |= 32;
            if (CheckBoxMuscleReplacement2.Checked)
                value2 |= 64;
            if (CheckBoxMuscleReplacement3.Checked)
                value2 |= 128;
            reader.WriteByte(Convert.ToByte(value2));

            reader.Close();
        }

        private void SaveGroupItems()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1 = 0, value2 = 0;

            // First set
            reader.Seek(0x1209C, SeekOrigin.Begin);
            if (checkBox15.Checked)
                value1 |= 1;
            if (checkBox14.Checked)
                value1 |= 2;
            if (checkBox13.Checked)
                value1 |= 4;
            if (checkBox12.Checked)
                value1 |= 8;
            if (checkBox11.Checked)
                value1 |= 16;
            if (checkBox10.Checked)
                value1 |= 32;
            if (checkBox9.Checked)
                value1 |= 64;
            reader.WriteByte(Convert.ToByte(value1));

            // Second set
            reader.Seek(0x1209D, SeekOrigin.Begin);
            if (checkBox8.Checked)
                value2 |= 1;
            if (checkBox7.Checked)
                value2 |= 2;
            if (checkBox6.Checked)
                value2 |= 4;
            if (checkBox5.Checked)
                value2 |= 8;
            if (checkBox4.Checked)
                value2 |= 16;
            if (checkBox3.Checked)
                value2 |= 32;
            if (checkBox2.Checked)
                value2 |= 64;
            if (checkBox1.Checked)
                value2 |= 128;
            reader.WriteByte(Convert.ToByte(value2));

            reader.Close();
        }

        private void SaveMatrixPasscodes()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            int value1 = 0, value2 = 0;

            // First set
            reader.Seek(0x120A0, SeekOrigin.Begin);
            if (checkBox19.Checked)
                value1 |= 1;
            if (checkBox18.Checked)
                value1 |= 2;
            if (checkBox17.Checked)
                value1 |= 4;
            if (checkBox16.Checked)
                value1 |= 8;
            reader.WriteByte(Convert.ToByte(value1));

            // Second set
            reader.Seek(0x120A1, SeekOrigin.Begin);
            if (checkBox31.Checked)
                value2 |= 1;
            if (checkBox32.Checked)
                value2 |= 2;
            if (checkBox34.Checked)
                value2 |= 4;
            if (checkBox35.Checked)
                value2 |= 8;
            if (checkBox33.Checked)
                value2 |= 16;
            if (checkBox29.Checked)
                value2 |= 32;
            if (checkBox30.Checked)
                value2 |= 64;
            if (checkBox28.Checked)
                value2 |= 128;
            reader.WriteByte(Convert.ToByte(value2));

            reader.Close();
        }

        private void LoadFile()
        {
            SetRunnerOffset();

            // Select Joshua as the default Spellbook
            ComboBoxSpellBook.SelectedIndex = 4;

            try
            {
                LoadBasics();
                LoadInventory();
                LoadAttributes();
                LoadSkills();
                LoadCyberdeck();
                LoadSpellbook();
                LoadCurrentRun();
                LoadCyberware();
                LoadGroupItems();
                LoadMatrixPasscodes();

                saveLoaded = true;
            }
            catch (Exception ex)
            {
                StringBuilder error = new StringBuilder();
                error.Append("Message:");
                error.AppendLine();
                error.Append(ex.Message);
                error.AppendLine();
                error.AppendLine();
                error.Append("Stack Trace:\n");
                error.AppendLine();
                error.Append(ex.StackTrace);

                FormError errorForm = new FormError();
                errorForm.Show();
                errorForm.TextBoxError.Text = error.ToString();

                if (reader != null)
                    reader.Close();
            }
        }

        private void SaveFile()
        {
            SetRunnerOffset();

            try
            {
                SaveBasics();
                SaveInventory();
                SaveAttributes();
                SaveSkills();
                SaveCyberdeck();
                SaveSpellbook();
                SaveCurrentRun();
                SaveCyberware();
                SaveGroupItems();
                SaveMatrixPasscodes();
            }
            catch (Exception ex)
            {
                StringBuilder error = new StringBuilder();
                error.Append("Message:");
                error.AppendLine();
                error.Append(ex.Message);
                error.AppendLine();
                error.AppendLine();
                error.Append("Stack Trace:\n");
                error.AppendLine();
                error.Append(ex.StackTrace);

                FormError errorForm = new FormError();
                errorForm.Show();
                errorForm.TextBoxError.Text = error.ToString();

                if (reader != null)
                    reader.Close();
            }
        }

        // Might use this eventually, if I find the offset
        private void ReadRunnerCount()
        {
            reader = new FileStream(TextBoxFile.Text, FileMode.Open);

            reader.Seek(0x1207E, SeekOrigin.Begin);
            runners = reader.ReadByte();

            if (runners == 0)
            {
                RadioButtonRunner1.Enabled = true;
            }
            if (runners == 1)
            {
                RadioButtonRunner1.Enabled = true;
                RadioButtonRunner2.Enabled = true;
            }
            if (runners == 2)
            {
                RadioButtonRunner1.Enabled = true;
                RadioButtonRunner2.Enabled = true;
                RadioButtonRunner3.Enabled = true;
            }

            reader.Close();
        }

        private void SetRunnerOffset()
        {
            if (RadioButtonRunner1.Checked)
                runnerOffset = runnerOffsets[0];
            if (RadioButtonRunner2.Checked)
                runnerOffset = runnerOffsets[1];
            if (RadioButtonRunner3.Checked)
                runnerOffset = runnerOffsets[2];
        }

        private void SetSpellOffset()
        {
            if (ComboBoxSpellBook.SelectedIndex == 0)
                spellOffset = spellOffsets[0];
            if (ComboBoxSpellBook.SelectedIndex == 1)
                spellOffset = spellOffsets[1];
            if (ComboBoxSpellBook.SelectedIndex == 2)
                spellOffset = spellOffsets[2];
            if (ComboBoxSpellBook.SelectedIndex == 3)
                spellOffset = spellOffsets[3];
            if (ComboBoxSpellBook.SelectedIndex == 4)
                spellOffset = spellOffsets[4];
        }

        private void ButtonLoad_Click(object sender, EventArgs e)
        {
            LoadFile();
        }

        private void ButtonSave_Click(object sender, EventArgs e)
        {
            SaveFile();
        }

        private void ButtonItemHelp_Click(object sender, EventArgs e)
        {
            StringBuilder helpString = new StringBuilder();
            helpString.Append("For items with a rating or quantity, just use a number\n\n");
            helpString.Append("For weapons, select the radio button next to the inventory slot, and right click on the Accessories Selector\n");
            helpString.Append("The value of the accessory will be added to the weapon\n");
            helpString.Append("If you mess up, reset the value to 0 or clear it with the accessory menu\n");

            MessageBox.Show(helpString.ToString(), "Inventory Mod & Accessory Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ButtonModifySpellsHelp_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Do not modify the total spells unless you know what you are doing! this number should be managed by the editor and is shown for testing purposes only.", "Total Spells", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CheckBoxModifySpells_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBoxModifySpells.Checked)
                NumericUpDownSpells.Enabled = true;
            else
                NumericUpDownSpells.Enabled = false;
        }

        private void ApplyItemImages()
        {
            List<PictureBox> pics = new List<PictureBox>();
            pics.Add(PictureBoxItem1);
            pics.Add(PictureBoxItem2);
            pics.Add(PictureBoxItem3);
            pics.Add(PictureBoxItem4);
            pics.Add(PictureBoxItem5);
            pics.Add(PictureBoxItem6);
            pics.Add(PictureBoxItem7);
            pics.Add(PictureBoxItem8);

            List<ComboBox> boxes = new List<ComboBox>();
            boxes.Add(ComboBoxInventory1);
            boxes.Add(ComboBoxInventory2);
            boxes.Add(ComboBoxInventory3);
            boxes.Add(ComboBoxInventory4);
            boxes.Add(ComboBoxInventory5);
            boxes.Add(ComboBoxInventory6);
            boxes.Add(ComboBoxInventory7);
            boxes.Add(ComboBoxInventory8);

            for (int i = 0; i < 8; i++)
                pics[i].Image = Image.FromFile(Application.StartupPath + @"\Items\" + boxes[i].Text + ".png");
        }

        private void ComboBoxInventory_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (saveLoaded)
                ApplyItemImages();
        }

        private void ButtonRaceHelp_Click(object sender, EventArgs e)
        {
            StringBuilder helpString = new StringBuilder();
            helpString.Append("Each race has it's own perks and flaws with regards to each attribute\n");
            helpString.Append("\n");
            helpString.Append("Human: Normal\n");
            helpString.Append("Elf: +1 Quickness, +2 Charisma\n");
            helpString.Append("Dwarf: +1 Body, -1 Quickness, +2 Strength, +1 Willpower\n");
            helpString.Append("Orc: +3 Body, +2 Strength, -1 Charisma, -1 Intelligence\n");
            helpString.Append("Troll: +5 Body, -1 Quickness, +4 Strength, -2 Charisma, -2 Intelligence, -1 Willpower\n");
            helpString.Append("Mage: No limits\n");

            MessageBox.Show(helpString.ToString(), "Race Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LabelVersion_Click(object sender, EventArgs e)
        {
            StringBuilder versionString = new StringBuilder();
            versionString.Append("Created by: Kyle873\n");
            versionString.Append("Email: Kyle873@gmail.com\n");
            versionString.Append("\n");
            versionString.Append("Now includes editing other runners' Spell Books!\n");
            versionString.Append("Matrix passcode bug fixed.\n");

            MessageBox.Show(versionString.ToString(), "Version Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LabelVersion_MouseEnter(object sender, EventArgs e)
        {
            Random rand = new Random();

            LabelVersion.ForeColor = Color.FromArgb(rand.Next(255), rand.Next(255), rand.Next(255));
        }

        private void LabelVersion_MouseLeave(object sender, EventArgs e)
        {
            LabelVersion.ForeColor = Color.Red;
        }

        private void ButtonResetSpells_Click(object sender, EventArgs e)
        {
            ComboBoxSpell1.SelectedIndex = 7;
            ComboBoxSpell2.SelectedIndex = 10;
            ComboBoxSpell3.SelectedIndex = 11;
            ComboBoxSpell4.SelectedIndex = 14;
            ComboBoxSpell5.SelectedIndex = 14;
            ComboBoxSpell6.SelectedIndex = 14;
            ComboBoxSpell7.SelectedIndex = 14;
            ComboBoxSpell8.SelectedIndex = 14;
            ComboBoxSpell9.SelectedIndex = 14;
            ComboBoxSpell10.SelectedIndex = 14;
            ComboBoxSpell11.SelectedIndex = 14;
            ComboBoxSpell12.SelectedIndex = 14;
            ComboBoxSpell13.SelectedIndex = 14;
            ComboBoxSpell14.SelectedIndex = 14;

            TrackBarSpellLevel1.Value = 3;
            TrackBarSpellLevel2.Value = 3;
            TrackBarSpellLevel3.Value = 1;
            TrackBarSpellLevel4.Value = 0;
            TrackBarSpellLevel5.Value = 0;
            TrackBarSpellLevel6.Value = 0;
            TrackBarSpellLevel7.Value = 0;
            TrackBarSpellLevel8.Value = 0;
            TrackBarSpellLevel9.Value = 0;
            TrackBarSpellLevel10.Value = 0;
            TrackBarSpellLevel11.Value = 0;
            TrackBarSpellLevel12.Value = 0;
            TrackBarSpellLevel13.Value = 0;
            TrackBarSpellLevel14.Value = 0;
        }

        private void ButtonGetAllSpells_Click(object sender, EventArgs e)
        {
            ComboBoxSpell1.SelectedIndex = 0;
            ComboBoxSpell2.SelectedIndex = 1;
            ComboBoxSpell3.SelectedIndex = 2;
            ComboBoxSpell4.SelectedIndex = 3;
            ComboBoxSpell5.SelectedIndex = 4;
            ComboBoxSpell6.SelectedIndex = 5;
            ComboBoxSpell7.SelectedIndex = 6;
            ComboBoxSpell8.SelectedIndex = 7;
            ComboBoxSpell9.SelectedIndex = 8;
            ComboBoxSpell10.SelectedIndex = 9;
            ComboBoxSpell11.SelectedIndex = 10;
            ComboBoxSpell12.SelectedIndex = 11;
            ComboBoxSpell13.SelectedIndex = 12;
            ComboBoxSpell14.SelectedIndex = 13;

            TrackBarSpellLevel1.Value = 8;
            TrackBarSpellLevel2.Value = 8;
            TrackBarSpellLevel3.Value = 8;
            TrackBarSpellLevel4.Value = 8;
            TrackBarSpellLevel5.Value = 8;
            TrackBarSpellLevel6.Value = 8;
            TrackBarSpellLevel7.Value = 8;
            TrackBarSpellLevel8.Value = 8;
            TrackBarSpellLevel9.Value = 8;
            TrackBarSpellLevel10.Value = 8;
            TrackBarSpellLevel11.Value = 8;
            TrackBarSpellLevel12.Value = 8;
            TrackBarSpellLevel13.Value = 8;
            TrackBarSpellLevel14.Value = 8;
        }

        private void ComboBoxRunType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Ghoul Bounty
            if (ComboBoxRunType.SelectedIndex == 0)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = false;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = false;
            }
            // Bodyguard
            if (ComboBoxRunType.SelectedIndex == 1)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = true;
                ComboBoxRunDestinationArea.Enabled = true;
                ComboBoxRunDestinationBuilding.Enabled = true;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = true;
            }
            // Courier
            if (ComboBoxRunType.SelectedIndex == 2)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = true;
                ComboBoxRunDestinationArea.Enabled = true;
                ComboBoxRunDestinationBuilding.Enabled = true;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = false;
            }
            // Enforcement
            if (ComboBoxRunType.SelectedIndex == 3)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = true;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = false;
            }
            // Acquisition
            if (ComboBoxRunType.SelectedIndex == 4)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = true;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = false;
            }
            // Extraction
            if (ComboBoxRunType.SelectedIndex == 5)
            {
                ComboBoxRunSourceArea.Enabled = true;
                ComboBoxRunSourceBuilding.Enabled = true;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = true;
            }
            // Matrix Run
            if (ComboBoxRunType.SelectedIndex == 6)
            {
                ComboBoxRunSourceArea.Enabled = false;
                ComboBoxRunSourceBuilding.Enabled = false;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = true;
                ComboBoxMatrixRunSystem.Enabled = true;
                ComboBoxMatrixRunListedSystem.Enabled = true;
                ComboBoxClientName.Enabled = false;
            }
            // None
            if (ComboBoxRunType.SelectedIndex == 7)
            {
                ComboBoxRunSourceArea.Enabled = false;
                ComboBoxRunSourceBuilding.Enabled = false;
                ComboBoxRunDestinationArea.Enabled = false;
                ComboBoxRunDestinationBuilding.Enabled = false;
                ComboBoxMatrixRunType.Enabled = false;
                ComboBoxMatrixRunSystem.Enabled = false;
                ComboBoxMatrixRunListedSystem.Enabled = false;
                ComboBoxClientName.Enabled = false;
            }
        }

        private void ToolStripMenuItemClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value = 0;
        }

        private void ToolStripMenuItemSmartlink_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 1;
        }

        private void ToolStripMenuItemSilencer_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 2;
        }

        private void ToolStripMenuItemSoundSupressor_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 4;
        }

        private void ToolStripMenuItemLaserSighting_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 8;
        }

        private void ToolStripMenuItemGasVent2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 16;
        }

        private void ToolStripMenuItemGasVent3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 8; i++)
                if (itemSlots[i].Checked)
                    itemValues[i].Value += 32;
        }

        private void ComboBoxRunSourceArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxRunSourceBuilding.Items.Clear();

            if (ComboBoxRunSourceArea.SelectedIndex == 0)
                ComboBoxRunSourceBuilding.Items.AddRange(area0);
            if (ComboBoxRunSourceArea.SelectedIndex == 1)
                ComboBoxRunSourceBuilding.Items.AddRange(area1);
            if (ComboBoxRunSourceArea.SelectedIndex == 2)
                ComboBoxRunSourceBuilding.Items.AddRange(area2);
            if (ComboBoxRunSourceArea.SelectedIndex == 3)
                ComboBoxRunSourceBuilding.Items.AddRange(area3);
            if (ComboBoxRunSourceArea.SelectedIndex == 4)
                ComboBoxRunSourceBuilding.Items.AddRange(area4);
            if (ComboBoxRunSourceArea.SelectedIndex == 5)
                ComboBoxRunSourceBuilding.Items.AddRange(area5);
            if (ComboBoxRunSourceArea.SelectedIndex == 6)
                ComboBoxRunSourceBuilding.Items.AddRange(area6);
        }

        private void ComboBoxRunDestinationArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBoxRunDestinationBuilding.Items.Clear();

            if (ComboBoxRunDestinationArea.SelectedIndex == 0)
                ComboBoxRunDestinationBuilding.Items.AddRange(area0);
            if (ComboBoxRunDestinationArea.SelectedIndex == 1)
                ComboBoxRunDestinationBuilding.Items.AddRange(area1);
            if (ComboBoxRunDestinationArea.SelectedIndex == 2)
                ComboBoxRunDestinationBuilding.Items.AddRange(area2);
            if (ComboBoxRunDestinationArea.SelectedIndex == 3)
                ComboBoxRunDestinationBuilding.Items.AddRange(area3);
            if (ComboBoxRunDestinationArea.SelectedIndex == 4)
                ComboBoxRunDestinationBuilding.Items.AddRange(area4);
            if (ComboBoxRunDestinationArea.SelectedIndex == 5)
                ComboBoxRunDestinationBuilding.Items.AddRange(area5);
            if (ComboBoxRunDestinationArea.SelectedIndex == 6)
                ComboBoxRunDestinationBuilding.Items.AddRange(area6);
        }

        private void ButtonClearRun_Click(object sender, EventArgs e)
        {
            ComboBoxCurrentJohnson.SelectedIndex = 5;
            ComboBoxRunType.SelectedIndex = 7;
            NumericUpDownPayment.Value = 0;
            ComboBoxRunSourceArea.SelectedIndex = 0;
            ComboBoxRunSourceBuilding.SelectedIndex = 0;
            ComboBoxRunDestinationArea.SelectedIndex = 0;
            ComboBoxRunDestinationBuilding.SelectedIndex = 0;
            ComboBoxMatrixRunType.SelectedIndex = 0;
            ComboBoxMatrixRunSystem.SelectedIndex = 0;
            ComboBoxMatrixRunListedSystem.SelectedIndex = 0;
            ComboBoxClientName.SelectedIndex = 0;
            NumericUpDownRunFlags.Value = 0;
        }

        private void LabelRunFlags_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not sure what this does yet, but it's here", "Flags", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RadioButtonRunner_CheckedChanged(object sender, EventArgs e)
        {
            if (saveLoaded)
            {
                SetRunnerOffset();
                LoadFile();
            }
        }

        private void ComboBoxSpellBook_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSpellOffset();
            LoadSpellbook();
        }
    }
}
