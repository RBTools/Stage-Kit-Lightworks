using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
using SK_Lightworks.StageKit;
using SlimDX.XInput;
using Controller = SlimDX.XInput.Controller;

namespace SK_Lightworks
{
    public partial class frmMain : Form
    {
        private Controller stageKitController;
        private StageKitController stageKit;
        private LedDisplay ledDisplay;
        private readonly Random random;
        private int led_delay;
        private bool showUpdateMessage;
        private const string AppName = "StageKit Lightworks";
        private bool ignorewait;
        private int globalwait;
        private readonly bool[] led_yellow_bool;
        private readonly bool[] led_blue_bool;
        private readonly bool[] led_red_bool;
        private readonly bool[] led_green_bool;
        private Thread scriptRunner;
        private string script_fullpath;
        private bool script_do_pause;
        private bool script_do_loop;
        private readonly Image LED_RedOn;
        private readonly Image LED_YellowOn;
        private readonly Image LED_GreenOn;
        private readonly Image LED_BlueOn;
        private readonly Image StageKitBackground;
        private readonly Image PowerOn;
        private readonly Image StrobeOn;
        private readonly Image FoggerOn;
        private readonly Tools MyTools;
        private readonly bool[] CurrentStateYellow;
        private readonly bool[] CurrentStateRed;
        private readonly bool[] CurrentStateGreen;
        private readonly bool[] CurrentStateBlue;
        private readonly PictureBox[] yellowBoxes;
        private readonly PictureBox[] blueBoxes;
        private readonly PictureBox[] greenBoxes;
        private readonly PictureBox[] redBoxes;
        private static readonly Color mMenuHighlight = Color.FromName("ControlDark");
        private static readonly Color mMenuBackground = Color.FromName("ControlDarkDark");
        private static readonly Color mMenuText = Color.WhiteSmoke;
        private static readonly Color mMenuBorder = Color.WhiteSmoke;
        private static readonly Color indicatorOFF = Color.Maroon;
        private static readonly Color indicatorON = Color.Lime;
        private int StrobeState;
        private bool FoggerIsOn;
        private bool doRedCircle;
        private bool doRedLaser;
        private bool doRedRandom;
        private int redIndex;
        private bool doBlueCircle;
        private bool doBlueLaser;
        private bool doBlueRandom;
        private int blueIndex;
        private bool doGreenCircle;
        private bool doGreenLaser;
        private bool doGreenRandom;
        private int greenIndex;
        private bool doYellowCircle;
        private bool doYellowLaser;
        private bool doYellowRandom;
        private int yellowIndex;

        private sealed class DarkRenderer : ToolStripProfessionalRenderer
        {
            public DarkRenderer() : base(new DarkColors()) { }
        }

        private sealed class DarkColors : ProfessionalColorTable
        {
            public override Color MenuItemSelected
            {
                get { return mMenuHighlight; }
            }
            public override Color MenuItemSelectedGradientBegin
            {
                get { return mMenuHighlight; }
            }
            public override Color MenuItemSelectedGradientEnd
            {
                get { return mMenuHighlight; }
            }
            public override Color MenuBorder
            {
                get { return mMenuBorder; }
            }
            public override Color MenuItemBorder
            {
                get { return mMenuBorder; }
            }
            public override Color MenuItemPressedGradientBegin
            {
                get { return mMenuHighlight; }
            }
            public override Color MenuItemPressedGradientEnd
            {
                get { return mMenuHighlight; }
            }
            public override Color MenuItemPressedGradientMiddle
            {
                get { return mMenuHighlight; }
            }
            public override Color CheckBackground
            {
                get { return mMenuHighlight; }
            }
            public override Color CheckPressedBackground
            {
                get { return mMenuHighlight; }
            }
            public override Color CheckSelectedBackground
            {
                get { return mMenuHighlight; }
            }
            public override Color ButtonSelectedBorder
            {
                get { return mMenuHighlight; }
            }
            public override Color SeparatorDark
            {
                get { return mMenuText; }
            }
            public override Color SeparatorLight
            {
                get { return mMenuText; }
            }
            public override Color ImageMarginGradientBegin
            {
                get { return mMenuBackground; }
            }
            public override Color ImageMarginGradientEnd
            {
                get { return mMenuBackground; }
            }
            public override Color ImageMarginGradientMiddle
            {
                get { return mMenuBackground; }
            }
            public override Color ToolStripDropDownBackground
            {
                get { return mMenuBackground; }
            }
        }

        public frmMain()
        {
            InitializeComponent();
            MyTools = new Tools();
            ledDisplay = new LedDisplay();
            stageKit = new StageKitController(1);
            random = new Random();
            menuStrip1.Renderer = new DarkRenderer();
            led_delay = 50;
            script_do_pause = false;
            script_do_loop = false;
            led_yellow_bool = new bool[8];
            led_red_bool = new bool[8];
            led_green_bool = new bool[8];
            led_blue_bool = new bool[8];
            CurrentStateBlue = new bool[8];
            CurrentStateGreen = new bool[8];
            CurrentStateRed = new bool[8];
            CurrentStateYellow = new bool[8];

            var path = Application.StartupPath + "\\res\\";
            LED_RedOn = MyTools.LoadImage(path + "red_on.png");
            LED_YellowOn = MyTools.LoadImage(path + "yellow_on.png");
            LED_GreenOn = MyTools.LoadImage(path + "green_on.png");
            LED_BlueOn = MyTools.LoadImage(path + "blue_on.png");
            StageKitBackground = MyTools.LoadImage(path + "stage_kit.png");
            PowerOn = MyTools.LoadImage(path + "power_on.png");
            StrobeOn = MyTools.LoadImage(path + "strobe_on.png");
            FoggerOn = MyTools.LoadImage(path + "fogger_on.png");

            yellowBoxes = new PictureBox[8];
            yellowBoxes[0] = led_yellow_led1;
            yellowBoxes[1] = led_yellow_led2;
            yellowBoxes[2] = led_yellow_led3;
            yellowBoxes[3] = led_yellow_led4;
            yellowBoxes[4] = led_yellow_led5;
            yellowBoxes[5] = led_yellow_led6;
            yellowBoxes[6] = led_yellow_led7;
            yellowBoxes[7] = led_yellow_led8;
            blueBoxes = new PictureBox[8];
            blueBoxes[0] = led_blue_led1;
            blueBoxes[1] = led_blue_led2;
            blueBoxes[2] = led_blue_led3;
            blueBoxes[3] = led_blue_led4;
            blueBoxes[4] = led_blue_led5;
            blueBoxes[5] = led_blue_led6;
            blueBoxes[6] = led_blue_led7;
            blueBoxes[7] = led_blue_led8;
            redBoxes = new PictureBox[8];
            redBoxes[0] = led_red_led1;
            redBoxes[1] = led_red_led2;
            redBoxes[2] = led_red_led3;
            redBoxes[3] = led_red_led4;
            redBoxes[4] = led_red_led5;
            redBoxes[5] = led_red_led6;
            redBoxes[6] = led_red_led7;
            redBoxes[7] = led_red_led8;
            greenBoxes = new PictureBox[8];
            greenBoxes[0] = led_green_led1;
            greenBoxes[1] = led_green_led2;
            greenBoxes[2] = led_green_led3;
            greenBoxes[3] = led_green_led4;
            greenBoxes[4] = led_green_led5;
            greenBoxes[5] = led_green_led6;
            greenBoxes[6] = led_green_led7;
            greenBoxes[7] = led_green_led8;

            for (var i = 0; i < 8; i++)
            {
                yellowBoxes[i] = AdjustImageParent(yellowBoxes[i]);
                blueBoxes[i] = AdjustImageParent(blueBoxes[i]);
                redBoxes[i] = AdjustImageParent(redBoxes[i]);
                greenBoxes[i] = AdjustImageParent(greenBoxes[i]);
            }
            AdjustImageParent(picPower);
            AdjustImageParent(picFogger);
            AdjustImageParent(picStrobe);
        }
        
        private PictureBox AdjustImageParent(Control image)
        {
            var pos = PointToScreen(image.Location);
            pos = picStageKit.PointToClient(pos);
            image.Parent = picStageKit;
            image.Location = pos;
            image.BackColor = Color.Transparent;
            return (PictureBox) image;
        }

        private static int GetControlTag(object sender)
        {
            int tag;
            try
            {
                tag = Convert.ToInt16((((Control) sender).Tag));
            }
            catch (Exception)
            {
                tag = 0;
            }
            return tag;
        }

        private void update_red_led_state(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var ledIndex = GetControlTag(sender);
            bool ledState;
            switch (ledIndex)
            {
                case 1:
                    stageKit.DisplayRedLed1(ref ledDisplay, !ledDisplay.RedLedArray.Led1);
                    ledState = ledDisplay.RedLedArray.Led1;
                    break;
                case 2:
                    stageKit.DisplayRedLed2(ref ledDisplay, !ledDisplay.RedLedArray.Led2);
                    ledState = ledDisplay.RedLedArray.Led2;
                    break;
                case 3:
                    stageKit.DisplayRedLed3(ref ledDisplay, !ledDisplay.RedLedArray.Led3);
                    ledState = ledDisplay.RedLedArray.Led3;
                    break;
                case 4:
                    stageKit.DisplayRedLed4(ref ledDisplay, !ledDisplay.RedLedArray.Led4);
                    ledState = ledDisplay.RedLedArray.Led4;
                    break;
                case 5:
                    stageKit.DisplayRedLed5(ref ledDisplay, !ledDisplay.RedLedArray.Led5);
                    ledState = ledDisplay.RedLedArray.Led5;
                    break;
                case 6:
                    stageKit.DisplayRedLed6(ref ledDisplay, !ledDisplay.RedLedArray.Led6);
                    ledState = ledDisplay.RedLedArray.Led6;
                    break;
                case 7:
                    stageKit.DisplayRedLed7(ref ledDisplay, !ledDisplay.RedLedArray.Led7);
                    ledState = ledDisplay.RedLedArray.Led7;
                    break;
                case 8:
                    stageKit.DisplayRedLed8(ref ledDisplay, !ledDisplay.RedLedArray.Led8);
                    ledState = ledDisplay.RedLedArray.Led8;
                    break;
                default:
                    return;
            }
            update_red_led_state(ledIndex - 1, ledState, ref ledDisplay, ref stageKit);
        }

        private void update_red_led_state(int ledIndex, bool ledState, ref LedDisplay display_panel, ref StageKitController controller_ref)
        {
            switch (ledIndex)
            {
                case 0:
                    controller_ref.DisplayRedLed1(ref display_panel, ledState);
                    break;
                case 1:
                    controller_ref.DisplayRedLed2(ref display_panel, ledState);
                    break;
                case 2:
                    controller_ref.DisplayRedLed3(ref display_panel, ledState);
                    break;
                case 3:
                    controller_ref.DisplayRedLed4(ref display_panel, ledState);
                    break;
                case 4:
                    controller_ref.DisplayRedLed5(ref display_panel, ledState);
                    break;
                case 5:
                    controller_ref.DisplayRedLed6(ref display_panel, ledState);
                    break;
                case 6:
                    controller_ref.DisplayRedLed7(ref display_panel, ledState);
                    break;
                case 7:
                    controller_ref.DisplayRedLed8(ref display_panel, ledState);
                    break;
                default:
                    return;
            }
            CurrentStateRed[ledIndex] = ledState;
            redBoxes[ledIndex].Image = ledState ? LED_RedOn : null;
        }

        private void update_yellow_led_state(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var ledIndex = GetControlTag(sender);
            bool ledState;
            switch (ledIndex)
            {
                case 1:
                    stageKit.DisplayYellowLed1(ref ledDisplay, !ledDisplay.YellowLedArray.Led1);
                    ledState = ledDisplay.YellowLedArray.Led1;
                    break;
                case 2:
                    stageKit.DisplayYellowLed2(ref ledDisplay, !ledDisplay.YellowLedArray.Led2);
                    ledState = ledDisplay.YellowLedArray.Led2;
                    break;
                case 3:
                    stageKit.DisplayYellowLed3(ref ledDisplay, !ledDisplay.YellowLedArray.Led3);
                    ledState = ledDisplay.YellowLedArray.Led3;
                    break;
                case 4:
                    stageKit.DisplayYellowLed4(ref ledDisplay, !ledDisplay.YellowLedArray.Led4);
                    ledState = ledDisplay.YellowLedArray.Led4;
                    break;
                case 5:
                    stageKit.DisplayYellowLed5(ref ledDisplay, !ledDisplay.YellowLedArray.Led5);
                    ledState = ledDisplay.YellowLedArray.Led5;
                    break;
                case 6:
                    stageKit.DisplayYellowLed6(ref ledDisplay, !ledDisplay.YellowLedArray.Led6);
                    ledState = ledDisplay.YellowLedArray.Led6;
                    break;
                case 7:
                    stageKit.DisplayYellowLed7(ref ledDisplay, !ledDisplay.YellowLedArray.Led7);
                    ledState = ledDisplay.YellowLedArray.Led7;
                    break;
                case 8:
                    stageKit.DisplayYellowLed8(ref ledDisplay, !ledDisplay.YellowLedArray.Led8);
                    ledState = ledDisplay.YellowLedArray.Led8;
                    break;
                default:
                    return;
            }
            update_yellow_led_state(ledIndex - 1, ledState, ref ledDisplay, ref stageKit);
        }

        private void update_yellow_led_state(int ledIndex, bool ledState, ref LedDisplay display_panel, ref StageKitController controller_ref)
        {
            switch (ledIndex)
            {
                case 0:
                    controller_ref.DisplayYellowLed1(ref display_panel, ledState);
                    break;
                case 1:
                    controller_ref.DisplayYellowLed2(ref display_panel, ledState);
                    break;
                case 2:
                    controller_ref.DisplayYellowLed3(ref display_panel, ledState);
                    break;
                case 3:
                    controller_ref.DisplayYellowLed4(ref display_panel, ledState);
                    break;
                case 4:
                    controller_ref.DisplayYellowLed5(ref display_panel, ledState);
                    break;
                case 5:
                    controller_ref.DisplayYellowLed6(ref display_panel, ledState);
                    break;
                case 6:
                    controller_ref.DisplayYellowLed7(ref display_panel, ledState);
                    break;
                case 7:
                    controller_ref.DisplayYellowLed8(ref display_panel, ledState);
                    break;
                default:
                    return;
            }
            CurrentStateYellow[ledIndex] = ledState;
            yellowBoxes[ledIndex].Image = ledState ? LED_YellowOn : null;
        }

        private void update_blue_led_state(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var ledIndex = GetControlTag(sender);
            bool ledState;
            switch (ledIndex)
            {
                case 1:
                    stageKit.DisplayBlueLed1(ref ledDisplay, !ledDisplay.BlueLedArray.Led1);
                    ledState = ledDisplay.BlueLedArray.Led1;
                    break;
                case 2:
                    stageKit.DisplayBlueLed2(ref ledDisplay, !ledDisplay.BlueLedArray.Led2);
                    ledState = ledDisplay.BlueLedArray.Led2;
                    break;
                case 3:
                    stageKit.DisplayBlueLed3(ref ledDisplay, !ledDisplay.BlueLedArray.Led3);
                    ledState = ledDisplay.BlueLedArray.Led3;
                    break;
                case 4:
                    stageKit.DisplayBlueLed4(ref ledDisplay, !ledDisplay.BlueLedArray.Led4);
                    ledState = ledDisplay.BlueLedArray.Led4;
                    break;
                case 5:
                    stageKit.DisplayBlueLed5(ref ledDisplay, !ledDisplay.BlueLedArray.Led5);
                    ledState = ledDisplay.BlueLedArray.Led5;
                    break;
                case 6:
                    stageKit.DisplayBlueLed6(ref ledDisplay, !ledDisplay.BlueLedArray.Led6);
                    ledState = ledDisplay.BlueLedArray.Led6;
                    break;
                case 7:
                    stageKit.DisplayBlueLed7(ref ledDisplay, !ledDisplay.BlueLedArray.Led7);
                    ledState = ledDisplay.BlueLedArray.Led7;
                    break;
                case 8:
                    stageKit.DisplayBlueLed8(ref ledDisplay, !ledDisplay.BlueLedArray.Led8);
                    ledState = ledDisplay.BlueLedArray.Led8;
                    break;
                default:
                    return;
            }
            update_blue_led_state(ledIndex - 1, ledState, ref ledDisplay, ref stageKit);
        }

        private void update_blue_led_state(int ledIndex, bool ledState, ref LedDisplay display_panel, ref StageKitController controller_ref)
        {
            switch (ledIndex)
            {
                case 0:
                    controller_ref.DisplayBlueLed1(ref display_panel, ledState);
                    break;
                case 1:
                    controller_ref.DisplayBlueLed2(ref display_panel, ledState);
                    break;
                case 2:
                    controller_ref.DisplayBlueLed3(ref display_panel, ledState);
                    break;
                case 3:
                    controller_ref.DisplayBlueLed4(ref display_panel, ledState);
                    break;
                case 4:
                    controller_ref.DisplayBlueLed5(ref display_panel, ledState);
                    break;
                case 5:
                    controller_ref.DisplayBlueLed6(ref display_panel, ledState);
                    break;
                case 6:
                    controller_ref.DisplayBlueLed7(ref display_panel, ledState);
                    break;
                case 7:
                    controller_ref.DisplayBlueLed8(ref display_panel, ledState);
                    break;
                default:
                    return;
            }
            CurrentStateBlue[ledIndex] = ledState;
            blueBoxes[ledIndex].Image = ledState ? LED_BlueOn : null;
        }

        private void update_green_led_state(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var ledIndex = GetControlTag(sender);
            bool ledState;
            switch (ledIndex)
            {
                case 1:
                    stageKit.DisplayGreenLed1(ref ledDisplay, !ledDisplay.GreenLedArray.Led1);
                    ledState = ledDisplay.GreenLedArray.Led1;
                    break;
                case 2:
                    stageKit.DisplayGreenLed2(ref ledDisplay, !ledDisplay.GreenLedArray.Led2);
                    ledState = ledDisplay.GreenLedArray.Led2;
                    break;
                case 3:
                    stageKit.DisplayGreenLed3(ref ledDisplay, !ledDisplay.GreenLedArray.Led3);
                    ledState = ledDisplay.GreenLedArray.Led3;
                    break;
                case 4:
                    stageKit.DisplayGreenLed4(ref ledDisplay, !ledDisplay.GreenLedArray.Led4);
                    ledState = ledDisplay.GreenLedArray.Led4;
                    break;
                case 5:
                    stageKit.DisplayGreenLed5(ref ledDisplay, !ledDisplay.GreenLedArray.Led5);
                    ledState = ledDisplay.GreenLedArray.Led5;
                    break;
                case 6:
                    stageKit.DisplayGreenLed6(ref ledDisplay, !ledDisplay.GreenLedArray.Led6);
                    ledState = ledDisplay.GreenLedArray.Led6;
                    break;
                case 7:
                    stageKit.DisplayGreenLed7(ref ledDisplay, !ledDisplay.GreenLedArray.Led7);
                    ledState = ledDisplay.GreenLedArray.Led7;
                    break;
                case 8:
                    stageKit.DisplayGreenLed8(ref ledDisplay, !ledDisplay.GreenLedArray.Led8);
                    ledState = ledDisplay.GreenLedArray.Led8;
                    break;
                default:
                    return;
            }
            update_green_led_state(ledIndex - 1, ledState, ref ledDisplay, ref stageKit);
        }

        private void update_green_led_state(int ledIndex, bool ledState, ref LedDisplay display_panel, ref StageKitController controller_ref)
        {
            switch (ledIndex)
            {
                case 0:
                    controller_ref.DisplayGreenLed1(ref display_panel, ledState);
                    break;
                case 1:
                    controller_ref.DisplayGreenLed2(ref display_panel, ledState);
                    break;
                case 2:
                    controller_ref.DisplayGreenLed3(ref display_panel, ledState);
                    break;
                case 3:
                    controller_ref.DisplayGreenLed4(ref display_panel, ledState);
                    break;
                case 4:
                    controller_ref.DisplayGreenLed5(ref display_panel, ledState);
                    break;
                case 5:
                    controller_ref.DisplayGreenLed6(ref display_panel, ledState);
                    break;
                case 6:
                    controller_ref.DisplayGreenLed7(ref display_panel, ledState);
                    break;
                case 7:
                    controller_ref.DisplayGreenLed8(ref display_panel, ledState);
                    break;
                default:
                    return;
            }
            CurrentStateGreen[ledIndex] = ledState;
            greenBoxes[ledIndex].Image = ledState ? LED_GreenOn : null;
        }
        
        private void UpdateStrobe()
        {
            if (StrobeState > 4)
            {
                StrobeState = 0;
            }
            strobe_slow_on.BackColor = indicatorOFF;
            strobe_medium_on.BackColor = indicatorOFF;
            strobe_fast_on.BackColor = indicatorOFF;
            strobe_faster_on.BackColor = indicatorOFF;
            StrobeSpeed speed;
            TextBox box;
            switch (StrobeState)
            {
                case 0:
                    stageKit.TurnStrobeOff();
                    return;
                case 1:
                    speed = StrobeSpeed.Slow;
                    box = strobe_slow_on;
                    break;
                case 2:
                    speed = StrobeSpeed.Medium;
                    box = strobe_medium_on;
                    break;
                case 3:
                    speed = StrobeSpeed.Faster;
                    box = strobe_fast_on;
                    break;
                case 4:
                    speed = StrobeSpeed.Fastest;
                    box = strobe_faster_on;
                    break;
                default:
                    return;
            }
            box.BackColor = indicatorON;
            stageKit.TurnStrobeOn(speed);
            strobeTimer.Enabled = !dontFlashStrobeLight.Checked;
        }

        private void StopTimers()
        {
            doRedCircle = false;
            doRedRandom = false;
            doRedLaser = false;
            doBlueCircle = false;
            doBlueRandom = false;
            doBlueLaser = false;
            doGreenCircle = false;
            doGreenRandom = false;
            doGreenLaser = false;
            doYellowCircle = false;
            doYellowRandom = false;
            doYellowLaser = false;
            redTimer.Enabled = false;
            blueTimer.Enabled = false;
            greenTimer.Enabled = false;
            yellowTimer.Enabled = false;
        }

        private void updateAllReds(bool state, bool updateButtons = false)
        {
            led_red_all_on.BackColor = state ? indicatorON : indicatorOFF;
            if (updateButtons)
            {
                redTimer.Enabled = false;
                doRedCircle = false;
                doRedRandom = false;
                doRedLaser = false;
                led_red_rand_on.BackColor = indicatorOFF;
                led_red_circle_on.BackColor = indicatorOFF;
                led_red_laser_on.BackColor = indicatorOFF;
                btnRedCircle.Tag = 0;
                btnRedLaser.Tag = 0;
                btnRedRandom.Tag = 0;
                redIndex = 0;
            }
            stageKit.DisplayRedAll(ref ledDisplay, state);
            for (var i = 0; i < 8; i++)
            {
                CurrentStateRed[i] = state;
                redBoxes[i].Image = state ? LED_RedOn : null;
            }
        }
        
        private void led_red_circle_Click(object sender, EventArgs e)
        {
            var button = (Button) sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllReds(false, true);
                return;
            }
            updateAllReds(false, true);
            button.Tag = 1;
            led_red_circle_on.BackColor = indicatorON;
            doRedCircle = true;
            redTimer.Enabled = true;
        }

        private void led_red_rand_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllReds(false, true);
                return;
            }
            updateAllReds(false, true);
            button.Tag = 1;
            led_red_rand_on.BackColor = indicatorON;
            doRedRandom = true;
            redTimer.Enabled = true;
        }

        private void led_red_laser_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllReds(false, true);
                return;
            }
            updateAllReds(false, true);
            button.Tag = 1;
            led_red_laser_on.BackColor = indicatorON;
            doRedLaser = true;
            redTimer.Enabled = true;
        }

        private void led_red_all_Click(object sender, EventArgs e)
        {
            updateAllReds(led_red_all_on.BackColor == indicatorOFF, true);
        }
        
        private void updateAllBlues(bool state, bool updateButtons = false)
        {
            led_blue_all_on.BackColor = state ? indicatorON : indicatorOFF;
            if (updateButtons)
            {
                blueTimer.Enabled = false;
                doBlueCircle = false;
                doBlueRandom = false;
                doBlueLaser = false;
                led_blue_rand_on.BackColor = indicatorOFF;
                led_blue_circle_on.BackColor = indicatorOFF;
                led_blue_laser_on.BackColor = indicatorOFF;
                btnBlueCircle.Tag = 0;
                btnBlueLaser.Tag = 0;
                btnBlueRandom.Tag = 0;
                blueIndex = 0;
            }
            stageKit.DisplayBlueAll(ref ledDisplay, state);
            for (var i = 0; i < 8; i++)
            {
                CurrentStateBlue[i] = state;
                blueBoxes[i].Image = state ? LED_BlueOn : null;
            }
        }

        private void led_blue_circle_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllBlues(false, true);
                return;
            }
            updateAllBlues(false, true);
            button.Tag = 1;
            led_blue_circle_on.BackColor = indicatorON;
            doBlueCircle = true;
            blueTimer.Enabled = true;
        }

        private void led_blue_rand_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllBlues(false, true);
                return;
            }
            updateAllBlues(false, true);
            button.Tag = 1;
            led_blue_rand_on.BackColor = indicatorON;
            doBlueRandom = true;
            blueTimer.Enabled = true;
        }

        private void led_blue_laser_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllBlues(false, true);
                return;
            }
            updateAllBlues(false, true);
            button.Tag = 1;
            led_blue_laser_on.BackColor = indicatorON;
            doBlueLaser = true;
            blueTimer.Enabled = true;
        }

        private void led_blue_all_Click(object sender, EventArgs e)
        {
            updateAllBlues(led_blue_all_on.BackColor == indicatorOFF, true);
        }
        
        private void updateAllGreens(bool state, bool updateButtons = false)
        {
            led_green_all_on.BackColor = state ? indicatorON : indicatorOFF;
            if (updateButtons)
            {
                greenTimer.Enabled = false;
                doGreenCircle = false;
                doGreenRandom = false;
                doGreenLaser = false;
                led_green_rand_on.BackColor = indicatorOFF;
                led_green_circle_on.BackColor = indicatorOFF;
                led_green_laser_on.BackColor = indicatorOFF;
                btnGreenCircle.Tag = 0;
                btnGreenLaser.Tag = 0;
                btnGreenRandom.Tag = 0;
                greenIndex = 0;
            }
            stageKit.DisplayGreenAll(ref ledDisplay, state);
            for (var i = 0; i < 8; i++)
            {
                CurrentStateGreen[i] = state;
                greenBoxes[i].Image = state ? LED_GreenOn : null;
            }
        }

        private void led_green_circle_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllGreens(false, true);
                return;
            }
            updateAllGreens(false, true);
            button.Tag = 1;
            led_green_circle_on.BackColor = indicatorON;
            doGreenCircle = true;
            greenTimer.Enabled = true;
        }

        private void led_green_rand_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllGreens(false, true);
                return;
            }
            updateAllGreens(false, true);
            button.Tag = 1;
            led_green_rand_on.BackColor = indicatorON;
            doGreenRandom = true;
            greenTimer.Enabled = true;
        }

        private void led_green_laser_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllGreens(false, true);
                return;
            }
            updateAllGreens(false, true);
            button.Tag = 1;
            led_green_laser_on.BackColor = indicatorON;
            doGreenLaser = true;
            greenTimer.Enabled = true;
        }

        private void led_green_all_Click(object sender, EventArgs e)
        {
            updateAllGreens(led_green_all_on.BackColor == indicatorOFF, true);
        }
        
        private void updateAllYellows(bool state, bool updateButtons = false)
        {
            led_yellow_all_on.BackColor = state ? indicatorON : indicatorOFF;
            if (updateButtons)
            {
                yellowTimer.Enabled = false;
                doYellowCircle = false;
                doYellowRandom = false;
                doYellowLaser = false;
                led_yellow_rand_on.BackColor = indicatorOFF;
                led_yellow_circle_on.BackColor = indicatorOFF;
                led_yellow_laser_on.BackColor = indicatorOFF;
                btnYellowCircle.Tag = 0;
                btnYellowLaser.Tag = 0;
                btnYellowRandom.Tag = 0;
                yellowIndex = 0;
            }
            stageKit.DisplayYellowAll(ref ledDisplay, state);
            for (var i = 0; i < 8; i++)
            {
                CurrentStateYellow[i] = state;
                yellowBoxes[i].Image = state ? LED_YellowOn : null;
            }
        }

        private void led_yellow_circle_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllYellows(false, true);
                return;
            }
            updateAllYellows(false, true);
            button.Tag = 1;
            led_yellow_circle_on.BackColor = indicatorON;
            doYellowCircle = true;
            yellowTimer.Enabled = true;
        }

        private void led_yellow_rand_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllYellows(false, true);
                return;
            }
            updateAllYellows(false, true);
            button.Tag = 1;
            led_yellow_rand_on.BackColor = indicatorON;
            doYellowRandom = true;
            yellowTimer.Enabled = true;
        }

        private void led_yellow_laser_Click(object sender, EventArgs e)
        {
            var button = (Button)sender;
            if (GetControlTag(sender) == 1)
            {
                updateAllYellows(false, true);
                return;
            }
            updateAllYellows(false, true);
            button.Tag = 1;
            led_yellow_laser_on.BackColor = indicatorON;
            doYellowLaser = true;
            yellowTimer.Enabled = true;
        }

        private void led_yellow_all_Click(object sender, EventArgs e)
        {
            updateAllYellows(led_yellow_all_on.BackColor == indicatorOFF, true);
        }
        
        private void led_reverse_circle_Click(object sender, EventArgs e)
        {
            btnRedCircle.PerformClick();
            btnBlueCircle.PerformClick();
            btnGreenCircle.PerformClick();
            btnYellowCircle.PerformClick();
        }

        private void led_reverse_rand_Click(object sender, EventArgs e)
        {
            btnRedRandom.PerformClick();
            btnBlueRandom.PerformClick();
            btnGreenRandom.PerformClick();
            btnYellowRandom.PerformClick();
        }

        private void led_reverse_laser_Click(object sender, EventArgs e)
        {
            btnRedLaser.PerformClick();
            btnBlueLaser.PerformClick();
            btnGreenLaser.PerformClick();
            btnYellowLaser.PerformClick();
        }

        private void led_reverse_all_Click(object sender, EventArgs e)
        {
            btnRedAllOn.PerformClick();
            btnBlueAllOn.PerformClick();
            btnGreenAllOn.PerformClick();
            btnYellowAllOn.PerformClick();
        }

        private void led_all_off_Click(object sender, EventArgs e)
        {
            StopLEDs();
        }

        private void kill_all_Click(object sender, EventArgs e)
        {
            StopEverything();
        }

        private void StopLEDs()
        {
            updateAllGreens(false, true);
            updateAllBlues(false, true);
            updateAllReds(false, true);
            updateAllYellows(false, true);
        }
        
        private void StopEverything(bool stopScript = true)
        {
            if (stopScript)
            {
                btnScriptStop.Invoke(new MethodInvoker(() => btnScriptStop.PerformClick()));
                btnScriptPause.Invoke(new MethodInvoker(() => btnScriptPause.Text = "Pause"));
            }
            StopLEDs();
            StopFogger();
            strobe_slow_on.BackColor = indicatorOFF;
            strobe_medium_on.BackColor = indicatorOFF;
            strobe_fast_on.BackColor = indicatorOFF;
            strobe_faster_on.BackColor = indicatorOFF;
            stageKit.TurnStrobeOff();
            strobeTimer.Enabled = false;
            picStrobe.Invoke(new MethodInvoker(() => picStrobe.Image = null));
            picFogger.Invoke(new MethodInvoker(() => picFogger.Image = null));
        }

        private void UpdateFogger(object sender, EventArgs e)
        {
            var duration = GetControlTag(sender);
            StopFogger();
            if (duration <= 0) return;
            UpdateFogger(duration * 1000);
        }

        private void UpdateFogger(int duration)
        {
            foggerTimer.Interval = duration;
            picFogger.Invoke(new MethodInvoker(() => picFogger.Image = FoggerOn));
            picFogger.Invoke(new MethodInvoker(() => picFogger.Refresh()));
            FoggerIsOn = true;
            stageKit.TurnFogOn();
            foggerTimer.Enabled = true;
        }

        private void StopFogger()
        {
            foggerTimer.Enabled = false;
            picFogger.Invoke(new MethodInvoker(() => picFogger.Image = null));
            picFogger.Invoke(new MethodInvoker(() => picFogger.Refresh()));
            FoggerIsOn = false;
            stageKit.TurnFogOff();
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopEverything();
        }

        private void connectionTimer_Tick(object sender, EventArgs e)
        {
            var connected = stageKitController.IsConnected;
            picPower.Image = connected ? PowerOn : null;
            toolTip1.SetToolTip(picPower, connected? "Controller is connected" : "Controller is not connected");
            if (!connected)
            {
                StopEverything();
                picStrobe.Image = null;
                picFogger.Image = null;
            }
            grpLED.Enabled = connected;
            grpStrobe.Enabled = connected;
            grpFogger.Enabled = connected;
            grpScripts.Enabled = connected;
            btnStopEverything.Enabled = connected;
            picStrobe.Enabled = connected;
            picFogger.Enabled = connected;
            strobeTimer.Enabled = !dontFlashStrobeLight.Checked && connected && stageKit.StrobeIsOn && StrobeState > 0;
            for (var i = 0; i < 8; i++)
            {
                blueBoxes[i].Enabled = connected;
                yellowBoxes[i].Enabled = connected;
                redBoxes[i].Enabled = connected;
                greenBoxes[i].Enabled = connected;
            }
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            Text = AppName;
            picStageKit.BackgroundImage = StageKitBackground;
            UpdateController(1);
            updater.RunWorkerAsync();
        }
        
        private void UpdateController(int index)
        {
            controllerOne.Checked = false;
            controllerTwo.Checked = false;
            controllerThree.Checked = false;
            controllerFour.Checked = false;
            var userIndex = UserIndex.One;
            switch (index)
            {
                default:
                    controllerOne.Checked = true;
                    break;
                case 2:
                    userIndex = UserIndex.Two;
                    controllerTwo.Checked = true;
                    break;
                case 3:
                    userIndex = UserIndex.Three;
                    controllerThree.Checked = true;
                    break;
                case 4:
                    userIndex = UserIndex.Four;
                    controllerFour.Checked = true;
                    break;
            }
            stageKitController = new Controller(userIndex);
            connectionTimer.Enabled = true;
        }

        private void controllerOne_Click(object sender, EventArgs e)
        {
            UpdateController(Convert.ToInt16(((ToolStripItem)sender).Tag));
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var version = GetAppVersion();
            var copyright = AppName + "\nVersion: " + version + "\nModified " + AppName + " © TrojanNemo, 2015\nOriginal SK Lightworks © Andrew Mussey, 2012" +
                            "\nStage Kit API © Brant Estes, 2009";
            const string message = "\n\nThis program is based on the unreleased beta source code found on Andrew Mussey's GitHub page for this project.\n" +
                                   "I took it, redesigned the GUI, heavily modified the source code, and I'm releasing it so others can find new uses for their " +
                                   "Rock Band Stage Kit.\n\nThis software is provided freely, but with no guarantees or assurances - report any bugs you find " +
                                   "and if and when I can update the software again, I'll try to take care of them.\n\nEnjoy.\n-TrojanNemo";
            MessageBox.Show(copyright + message, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string GetAppVersion()
        {
            var vers = Assembly.GetExecutingAssembly().GetName().Version;
            return "v" + String.Format("{0}.{1}.{2}", vers.Major, vers.Minor, vers.Build);
        }

        private void led_speed_slower_Click(object sender, EventArgs e)
        {
            ChangeLEDSpeed(false);
        }

        private void led_speed_faster_Click(object sender, EventArgs e)
        {
            ChangeLEDSpeed(true);
        }

        private void ChangeLEDSpeed(bool faster)
        {
            if (faster)
            {
                led_delay += 5;
            }
            else
            {
                led_delay -= 5;
            }
            UpdateLEDSpeed();
        }

        private void UpdateLEDSpeed()
        {
            if (led_delay < 0)
            {
                led_delay = 0;
            }
            if (led_delay > 999)
            {
                led_delay = 999;
            }
            UpdateTimerIntervalSpeed();
            txtDelay.Text = "Delay: " + led_delay;
        }

        private void UpdateTimerIntervalSpeed()
        {
            redTimer.Interval = led_delay;
            blueTimer.Interval = led_delay;
            greenTimer.Interval = led_delay;
            yellowTimer.Interval = led_delay;
        }

        private void led_speed_box_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData != Keys.Return) return;
            try
            {
                led_delay = Convert.ToInt16(txtDelay.Text.Replace("Delay:","").Trim());
            }
            catch (Exception)
            {
                MessageBox.Show("That's not a valid delay value, try again", AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            UpdateLEDSpeed();
        }

        private void updater_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var path = Application.StartupPath + "\\bin\\update.txt";
            MyTools.DeleteFile(path);
            using (var client = new WebClient())
            {
                try
                {
                    client.DownloadFile("http://www.keepitfishy.com/rb3/stagekit/update.txt", path);
                }
                catch (Exception)
                { }
            }
        }

        private void updater_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            var path = Application.StartupPath + "\\bin\\update.txt";
            if (!File.Exists(path))
            {
                if (showUpdateMessage)
                {
                    MessageBox.Show("Unable to check for updates", AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                return;
            }
            var thisVersion = GetAppVersion();
            var newVersion = "v";
            string newName;
            string releaseDate;
            string link;
            var changeLog = new List<string>();
            var sr = new StreamReader(path);
            try
            {
                var line = sr.ReadLine();
                if (line.ToLowerInvariant().Contains("html"))
                {
                    sr.Dispose();
                    if (showUpdateMessage)
                    {
                        MessageBox.Show("Unable to check for updates", AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    return;
                }
                newName = MyTools.GetConfigString(line);
                newVersion += MyTools.GetConfigString(sr.ReadLine());
                releaseDate = MyTools.GetConfigString(sr.ReadLine());
                link = MyTools.GetConfigString(sr.ReadLine());
                sr.ReadLine();//ignore Change Log header
                while (sr.Peek() >= 0)
                {
                    changeLog.Add(sr.ReadLine());
                }
            }
            catch (Exception ex)
            {
                if (showUpdateMessage)
                {
                    MessageBox.Show("Error parsing update file:\n" + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                sr.Dispose();
                return;
            }
            sr.Dispose();
            MyTools.DeleteFile(path);
            if (thisVersion.Equals(newVersion))
            {
                if (showUpdateMessage)
                {
                    MessageBox.Show("You have the latest version", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }
            var newInt = Convert.ToInt16(newVersion.Replace("v", "").Replace(".", "").Trim());
            var thisInt = Convert.ToInt16(thisVersion.Replace("v", "").Replace(".", "").Trim());
            if (newInt <= thisInt)
            {
                if (showUpdateMessage)
                {
                    MessageBox.Show("You have a newer version (" + thisVersion + ") than what's on the server (" + newVersion + ")\nNo update needed!", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }
            var updaterForm = new Updater();
            updaterForm.SetInfo(AppName, thisVersion, newName, newVersion, releaseDate, link, changeLog);
            updaterForm.ShowDialog();
        }

        private void checkForUpdates_Click(object sender, EventArgs e)
        {
            showUpdateMessage = true;
            updater.RunWorkerAsync();
        }

        private void script_run_Click(object sender, EventArgs e)
        {
            btnScriptStop.Enabled = true;
            btnScriptRun.Enabled = false;
            btnScriptLoad.Enabled = false;
            btnScriptPause.Enabled = true;
            scriptRunner = new Thread(script_run_helper);
            scriptRunner.Start();
        }

        private void script_loop_Click(object sender, EventArgs e)
        {
            script_do_loop = !script_do_loop;
            loop_on.BackColor = script_do_loop ? indicatorON : indicatorOFF;
        }

        private void script_stop_Click(object sender, EventArgs e)
        {
            script_do_pause = false;
            scriptRunner.Abort();
            btnScriptRun.Enabled = true;
            btnScriptStop.Enabled = false;
            btnScriptPause.Enabled = false;
            btnScriptLoad.Enabled = true;
            btnScriptPause.Text = "Pause";
        }
        
        private void script_run_helper()
        {
            if (script_do_pause) return;
            var command = "";
            var lineNumber = 0;
            do
            {
                var sr = new StreamReader(script_fullpath);
                try
                {
                    globalwait = 0;
                    while ((command = sr.ReadLine()) != null)
                    {
                        while (script_do_pause)
                        {
                            //just hang out here until unpaused
                        }
                        lineNumber++;
                        script_run_command_helper(command, lineNumber);
                    }
                    sr.Dispose();
                }
                catch (ThreadAbortException)
                {
                    sr.Dispose();
                    try
                    {
                        txtScriptCommand.Invoke(new MethodInvoker(() => txtScriptCommand.Text = " Command: None"));
                        btnScriptStop.Invoke(new MethodInvoker(() => btnScriptStop.PerformClick()));
                        btnStopEverything.Invoke(new MethodInvoker(() => btnStopEverything.PerformClick()));
                    }
                    catch (Exception)
                    {} //will throw exception if called when disposing form, just ignore
                    
                }
                catch (Exception ex)
                {
                    sr.Dispose();
                    MessageBox.Show("Error processing script command '" + command + "'\nError message: " + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                }
            } 
            while (script_do_loop);
            txtScriptCommand.Invoke(new MethodInvoker(() => txtScriptCommand.Text = " Command: None"));
            btnScriptStop.Invoke(new MethodInvoker(() => btnScriptStop.PerformClick()));
            btnStopEverything.Invoke(new MethodInvoker(() => btnStopEverything.PerformClick()));
        }
        
        private void script_run_command_helper(string command, int lineNumber)
        {
            txtScriptCommand.Invoke(new MethodInvoker(() => txtScriptCommand.Text = " Command: " + command));
            var commands = command.Trim().Split(' ');
            if (!commands.Any()) return;
            int index;
            ignorewait = false;
            switch (commands[0].ToLowerInvariant())
            {
                case "killall":
                    StopEverything(false);
                    break;
                case "blue":
                    index = Convert.ToInt16(commands[1]) - 1;
                    led_blue_bool[index] = commands.Length <= 2 ? !led_blue_bool[index] : commands[2].Equals("1") || !commands[2].Equals("0") && !led_blue_bool[index];
                    update_blue_led_state(index, led_blue_bool[index], ref ledDisplay, ref stageKit);
                    break;
                case "blueall":
                    for (var i = 0; i < 8; i++)
                    {
                        if (commands.Length == 1)
                        {
                            led_blue_bool[i] = !led_blue_bool[i];
                        }
                        else
                        {
                            led_blue_bool[i] = commands[1] == "1";
                        }
                        update_blue_led_state(i, led_blue_bool[i], ref ledDisplay, ref stageKit);
                    }
                    break;
                case "green":
                    index = Convert.ToInt16(commands[1]) - 1;
                    led_green_bool[index] = commands.Length <= 2 ? !led_green_bool[index] : commands[2].Equals("1") || !commands[2].Equals("0") && !led_green_bool[index];
                    update_green_led_state(index, led_green_bool[index], ref ledDisplay, ref stageKit);
                    break;
                case "greenall":
                    for (var i = 0; i < 8; i++)
                    {
                        if (commands.Length == 1)
                        {
                            led_green_bool[i] = !led_green_bool[i];
                        }
                        else
                        {
                            led_green_bool[i] = commands[1] == "1";
                        }
                        update_green_led_state(i, led_green_bool[i], ref ledDisplay, ref stageKit);
                    }
                    break;
                case "red":
                    index = Convert.ToInt16(commands[1]) - 1;
                    led_red_bool[index] = commands.Length <= 2 ? !led_red_bool[index] : commands[2].Equals("1") || !commands[2].Equals("0") && !led_red_bool[index];
                    update_red_led_state(index, led_red_bool[index], ref ledDisplay, ref stageKit);
                    break;
                case "redall":
                    for (var i = 0; i < 8; i++)
                    {
                        if (commands.Length == 1)
                        {
                            led_red_bool[i] = !led_red_bool[i];
                        }
                        else
                        {
                            led_red_bool[i] = commands[1] == "1";
                        }
                        update_red_led_state(i, led_red_bool[i], ref ledDisplay, ref stageKit);
                    }
                    break;
                case "yellow":
                    index = Convert.ToInt16(commands[1]) - 1;
                    led_yellow_bool[index] = commands.Length <= 2 ? !led_yellow_bool[index] : commands[2].Equals("1") || !commands[2].Equals("0") && !led_yellow_bool[index];
                    update_yellow_led_state(index, led_yellow_bool[index], ref ledDisplay, ref stageKit);
                    break;
                case "yellowall":
                    for (var i = 0; i < 8; i++)
                    {
                        if (commands.Length == 1)
                        {
                            led_yellow_bool[i] = !led_yellow_bool[i];
                        }
                        else
                        {
                            led_yellow_bool[i] = commands[1] == "1";
                        }
                        update_yellow_led_state(i, led_yellow_bool[i], ref ledDisplay, ref stageKit);
                    }
                    break;
                case "strobe":
                    StrobeState = Convert.ToInt16(commands[1]);
                    UpdateStrobe();
                    break;
                case "fog":
                    UpdateFogger(Convert.ToInt32(commands[1]));
                    break;
                case "wait":
                    Thread.Sleep(Convert.ToInt32(commands[1]));
                    break;
                case "globalwait":
                    globalwait = Convert.ToInt32(commands[1]);
                    ignorewait = true;
                    break;
                default:
                    ignorewait = true;
                    if (!string.IsNullOrEmpty(commands[0].Trim()) && !commands[0].StartsWith("#", StringComparison.Ordinal))
                    {
                        MessageBox.Show("Command '" + command + "' at line number " + lineNumber + " is not a valid script command", AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    break;
            }
            if (ignorewait) return;
            Thread.Sleep(globalwait);
        }
        
        private void script_pause_Click(object sender, EventArgs e)
        {
            script_do_pause = !script_do_pause;
            btnScriptPause.Text = script_do_pause ? "Resume" : "Pause";
            btnScriptRun.Enabled = false;
            btnScriptStop.Enabled = true;
            btnScriptLoad.Enabled = false;
            btnScriptPause.Enabled = true;
            if (script_do_pause)
            {
                txtScriptCommand.Invoke(new MethodInvoker(() => txtScriptCommand.Text = " [PAUSED]"));
            }
        }

        private void script_load_Click(object sender, EventArgs e)
        {
            var path = Application.StartupPath + "\\scripts\\";
            var ofd = new OpenFileDialog
            {
                InitialDirectory = Directory.Exists(path) ? path : Environment.CurrentDirectory,
                Filter = "SK Lightworks Script files (*.sk)|*.sk",
                Multiselect = false,
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            Environment.CurrentDirectory = Path.GetDirectoryName(ofd.FileName);
            try
            {
                txtScript.Text = " Script: " + ofd.SafeFileName;
                script_fullpath = ofd.FileName;
                btnScriptRun.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading script file:\n" + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        
        private void led_speed_box_MouseClick(object sender, MouseEventArgs e)
        {
            txtDelay.Text = led_delay.ToString(CultureInfo.InvariantCulture);
            txtDelay.SelectionStart = txtDelay.TextLength;
        }

        private void led_speed_box_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Return) return;
            e.Handled = true;
        }

        private void UpdateStrobe(object sender, EventArgs e)
        {
            StrobeState = GetControlTag(sender);
            UpdateStrobe();
        }

        private void picStrobe_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            StrobeState++;
            UpdateStrobe();
        }

        private void picFogger_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            UpdateFogger(FoggerIsOn ? btnFogOff : btnFogOn, null);
        }

        private void strobeTimer_Tick(object sender, EventArgs e)
        {
            if (StrobeState == 0 || !stageKit.StrobeIsOn)
            {
                strobeTimer.Enabled = false;
                picStrobe.Invoke(new MethodInvoker(() => picStrobe.Image = null));
                return;
            }
            switch (StrobeState)
            {
                case 1:
                    strobeTimer.Interval = 100;
                    break;
                case 2:
                    strobeTimer.Interval = 75;
                    break;
                case 3:
                    strobeTimer.Interval = 50;
                    break;
                case 4:
                    strobeTimer.Interval = 25;
                    break;
            }
            if (strobeTimer.Tag.ToString() == "1")
            {
                picStrobe.Invoke(new MethodInvoker(() => picStrobe.Image = null));
                strobeTimer.Tag = 0;
            }
            else
            {
                picStrobe.Invoke(new MethodInvoker(() => picStrobe.Image = dontFlashStrobeLight.Checked ? null : StrobeOn));
                strobeTimer.Tag = 1;
            }
        }

        private void dontFlashStrobeLight_Click(object sender, EventArgs e)
        {
            if (!dontFlashStrobeLight.Checked) return;
            picStrobe.Image = null;
            strobeTimer.Enabled = false;
        }

        private void c3Forums_Click(object sender, EventArgs e)
        {
            Process.Start("http://customscreators.com/index.php?/topic/13378-stagekit-lightworks-v300-9415-use-your-stage-kit-on-your-pc/");
        }
        
        private void foggerTimer_Tick(object sender, EventArgs e)
        {
            stageKit.TurnFogOff();
            foggerTimer.Enabled = false;
            picFogger.Image = null;
            FoggerIsOn = false;
        }
        
        private void redTimer_Tick(object sender, EventArgs e)
        {
            if (doRedCircle)
            {
                redIndex++;
                if (redIndex > 7)
                {
                    redIndex = 0;
                }
                update_red_led_state(redIndex, true, ref ledDisplay, ref stageKit);
                update_red_led_state(redIndex == 0 ? 7 : redIndex - 1, false, ref ledDisplay, ref stageKit);
            }
            else if (doRedLaser)
            {
                update_red_led_state(redIndex, false, ref ledDisplay, ref stageKit);
                redIndex = random.Next(0, 8);
                update_red_led_state(redIndex, true, ref ledDisplay, ref stageKit);
            }
            else if (doRedRandom)
            {
                var index = random.Next(0, 8);
                update_red_led_state(index, !CurrentStateRed[index], ref ledDisplay, ref stageKit);
            }
            else
            {
                updateAllReds(false);
                redTimer.Enabled = false;
            }
        }

        private void blueTimer_Tick(object sender, EventArgs e)
        {
            if (doBlueCircle)
            {
                blueIndex++;
                if (blueIndex > 7)
                {
                    blueIndex = 0;
                }
                update_blue_led_state(blueIndex, true, ref ledDisplay, ref stageKit);
                update_blue_led_state(blueIndex == 0 ? 7 : blueIndex - 1, false, ref ledDisplay, ref stageKit);
            }
            else if (doBlueLaser)
            {
                update_blue_led_state(blueIndex, false, ref ledDisplay, ref stageKit);
                blueIndex = random.Next(0, 8);
                update_blue_led_state(blueIndex, true, ref ledDisplay, ref stageKit);
            }
            else if (doBlueRandom)
            {
                var index = random.Next(0, 8);
                update_blue_led_state(index, !CurrentStateBlue[index], ref ledDisplay, ref stageKit);
            }
            else
            {
                updateAllBlues(false);
                blueTimer.Enabled = false;
            }
        }

        private void greenTimer_Tick(object sender, EventArgs e)
        {
            if (doGreenCircle)
            {
                greenIndex++;
                if (greenIndex > 7)
                {
                    greenIndex = 0;
                }
                update_green_led_state(greenIndex, true, ref ledDisplay, ref stageKit);
                update_green_led_state(greenIndex == 0 ? 7 : greenIndex - 1, false, ref ledDisplay, ref stageKit);
            }
            else if (doGreenLaser)
            {
                update_green_led_state(greenIndex, false, ref ledDisplay, ref stageKit);
                greenIndex = random.Next(0, 8);
                update_green_led_state(greenIndex, true, ref ledDisplay, ref stageKit);
            }
            else if (doGreenRandom)
            {
                var index = random.Next(0, 8);
                update_green_led_state(index, !CurrentStateGreen[index], ref ledDisplay, ref stageKit);
            }
            else
            {
                updateAllGreens(false);
                greenTimer.Enabled = false;
            }
        }

        private void yellowTimer_Tick(object sender, EventArgs e)
        {
            if (doYellowCircle)
            {
                yellowIndex++;
                if (yellowIndex > 7)
                {
                    yellowIndex = 0;
                }
                update_yellow_led_state(yellowIndex, true, ref ledDisplay, ref stageKit);
                update_yellow_led_state(yellowIndex == 0 ? 7 : yellowIndex - 1, false, ref ledDisplay, ref stageKit);
            }
            else if (doYellowLaser)
            {
                update_yellow_led_state(yellowIndex, false, ref ledDisplay, ref stageKit);
                yellowIndex = random.Next(0, 8);
                update_yellow_led_state(yellowIndex, true, ref ledDisplay, ref stageKit);
            }
            else if (doYellowRandom)
            {
                var index = random.Next(0, 8);
                update_yellow_led_state(index, !CurrentStateYellow[index], ref ledDisplay, ref stageKit);
            }
            else
            {
                updateAllYellows(false);
                yellowTimer.Enabled = false;
            }
        }
    }

    public class Tools
    {
        public void DeleteFile(string file)
        {
            if (string.IsNullOrEmpty(file) || !File.Exists(file)) return;
            try
            {
                File.Delete(file);
            }
            catch (Exception)
            { }
        }

        public string GetConfigString(string raw_line)
        {
            var line = raw_line;
            var index = line.IndexOf("=", StringComparison.Ordinal) + 1;
            try
            {
                line = line.Substring(index, line.Length - index);
            }
            catch (Exception)
            {
                line = "";
            }
            return line.Trim();
        }

        public Image LoadImage(string file)
        {
            if (!File.Exists(file)) return null;
            Image img;
            using (var bmpTemp = new Bitmap(file))
            {
                img = new Bitmap(bmpTemp);
            }
            return img;
        }
    }
}
