// Version 1.0.0 2023-02-28
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;

using Autodesk.Revit.UI;

namespace BuildRibbonFromXML //.Ribbon_Builder
{
    public class RibbonBuilder
    {
        public static void build_ribbon(UIControlledApplication application, string input_assembly_file_path, string input_ribbon_file_path) // Using UI Controlled Application for the ability to add items to the UI
        {
            // Assign variable for where implementation code is
            string this_assembly_path = input_assembly_file_path;

            // Assign variable for where xml file is
            string ribbon_xml_file_path = input_ribbon_file_path;
            string[] ribbon_xml_file = System.IO.Directory.GetFiles(ribbon_xml_file_path, "*.ribbon");

            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(ribbon_xml_file[0]); // Set location of xml file

            XmlNodeList tab_list = xdoc.SelectNodes("tab");

            foreach (XmlNode tab in tab_list) // For each tab 

            {
                string tab_name_string = tab.Attributes["name"].Value;
                application.CreateRibbonTab(tab_name_string);

                XmlNodeList panel_list = tab.SelectNodes("panel");

                foreach (XmlNode panel in panel_list) // For each panel
                {
                    string panel_name_string = panel.Attributes["name"].Value;
                    RibbonPanel ribbon_panel = application.CreateRibbonPanel(tab_name_string, panel_name_string);

                    XmlNodeList button_structure_list = panel.ChildNodes;

                    foreach (XmlNode button_structure in button_structure_list) // For each button or object
                    {
                        if (button_structure.Name == "separator")
                        {
                            ribbon_panel.AddSeparator();
                        }

                        if (button_structure.Name == "stackeditems")
                        {
                            XmlNodeList item_list = button_structure.ChildNodes;
                            int item_list_count = item_list.Count;

                            if (item_list_count == 0 | item_list_count > 3) // if there are no items or more than three, exit if statement
                            {
                                break;
                            }

                            if (item_list_count == 1) // for one item
                            {
                                foreach (XmlNode item in item_list)
                                {
                                    if (item.Name == "button")
                                    {
                                        ribbon_panel.AddItem(add_button(item, this_assembly_path));
                                    }

                                    if (item.Name == "pulldownbuttons")
                                    {
                                        PulldownButton pulldown_button = ribbon_panel.AddItem(add_pulldown(item)) as PulldownButton;

                                        XmlNodeList button_list = item.ChildNodes;

                                        foreach (XmlNode button in button_list)
                                        {
                                            pulldown_button.AddPushButton(add_button(button, this_assembly_path));
                                        }
                                    }

                                    if (item.Name == "combobox")
                                    {
                                        string combobox_name_string = item.Attributes["name"].Value;
                                        ComboBoxData combobox_data = new ComboBoxData(combobox_name_string);
                                        ComboBox combobox = ribbon_panel.AddItem(combobox_data) as ComboBox;

                                        add_combobox(item, combobox);

                                        // Use this Event Handler Template for a Combo Box

                                        //combobox.CurrentChanged += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs>(combo_box_event);
                                        //combobox.DropDownClosed += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownClosedEventArgs>(combo_box_event);
                                        //combobox.DropDownOpened += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs>(combo_box_event);

                                    }

                                    if (item.Name == "textbox")
                                    {
                                        TextBoxData textbox_data = add_textbox(item);
                                        TextBox textbox = ribbon_panel.AddItem(textbox_data) as TextBox;

                                        if (item.Attributes["prompttext"].Value != "")
                                        {
                                            textbox.PromptText = item.Attributes["prompttext"].Value;
                                        }

                                        if (item.Attributes["showimage"].Value == "true")
                                        {
                                            textbox.ShowImageAsButton = true;
                                        }

                                        //// Use this Event Handler Template for a Text Box

                                        // textbox.EnterPressed += new EventHandler<Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs>(text_box_event);
                                    }
                                }
                            }

                            List<RibbonItem> ribbon_stacked_items = new List<RibbonItem>();
                            List<RibbonItemData> item_data_list = new List<RibbonItemData>();


                            if (item_list_count == 2 || item_list_count == 3) // for two or three items
                            {
                                foreach (XmlNode item in item_list)
                                {
                                    if (item.Name == "button")
                                    {
                                        item_data_list.Add(add_button(item, this_assembly_path));
                                    }

                                    if (item.Name == "pulldownbuttons")
                                    {
                                        item_data_list.Add(add_pulldown(item));
                                    }

                                    if (item.Name == "combobox")
                                    {
                                        string combobox_name_string = item.Attributes["name"].Value;
                                        ComboBoxData combobox_data = new ComboBoxData(combobox_name_string);
                                        item_data_list.Add(combobox_data);
                                    }

                                    if (item.Name == "textbox")
                                    {
                                        TextBoxData textbox_data = add_textbox(item);
                                        item_data_list.Add(textbox_data);
                                    }

                                }

                                if (item_list_count == 2)
                                {
                                    ribbon_stacked_items.AddRange(ribbon_panel.AddStackedItems(item_data_list[0], item_data_list[1]));

                                    for (int i = 0; i < 2; i++)
                                    {
                                        if (ribbon_stacked_items[i].ItemType.ToString() == "PulldownButton")
                                        {
                                            XmlNodeList button_list = item_list[i].ChildNodes;

                                            foreach (XmlNode button in button_list)
                                            {
                                                ((PulldownButton)ribbon_stacked_items[i]).AddPushButton(add_button(button, this_assembly_path));
                                            }
                                        }

                                        if (ribbon_stacked_items[i].ItemType.ToString() == "ComboBox")
                                        {
                                            add_combobox(item_list[i], (ComboBox)ribbon_stacked_items[i]);

                                            //// Use this Event Handler Template for a Combo Box

                                            //((ComboBox)ribbon_stacked_items[i]).CurrentChanged += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs>(combo_box_event);
                                            //((ComboBox)ribbon_stacked_items[i]).DropDownClosed += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownClosedEventArgs>(combo_box_event);
                                            //((ComboBox)ribbon_stacked_items[i]).DropDownOpened += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs>(combo_box_event);
                                        }

                                        if (ribbon_stacked_items[i].ItemType.ToString() == "TextBox")
                                        {
                                            if (item_list[i].Attributes["prompttext"].Value != "")
                                            {
                                                ((TextBox)ribbon_stacked_items[i]).PromptText = item_list[i].Attributes["prompttext"].Value;
                                            }

                                            if (item_list[i].Attributes["showimage"].Value == "true")
                                            {
                                                ((TextBox)ribbon_stacked_items[i]).ShowImageAsButton = true;
                                            }

                                            // Use this Event Handler Template for a Text Box

                                            // textbox.EnterPressed += new EventHandler<Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs>(text_box_event);
                                        }
                                    }
                                }

                                if (item_list_count == 3)
                                {
                                    ribbon_stacked_items.AddRange(ribbon_panel.AddStackedItems(item_data_list[0], item_data_list[1], item_data_list[2]));

                                    for (int i = 0; i < 3; i++)
                                    {
                                        if (ribbon_stacked_items[i].ItemType.ToString() == "PulldownButton")
                                        {
                                            XmlNodeList button_list = item_list[i].ChildNodes;

                                            foreach (XmlNode button in button_list)
                                            {
                                                ((PulldownButton)ribbon_stacked_items[i]).AddPushButton(add_button(button, this_assembly_path));
                                            }
                                        }

                                        if (ribbon_stacked_items[i].ItemType.ToString() == "ComboBox")
                                        {
                                            add_combobox(item_list[i], (ComboBox)ribbon_stacked_items[i]);

                                            //// Use this Event Handler Template for a Text Box

                                            //((ComboBox)ribbon_stacked_items[i]).CurrentChanged += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs>(combo_box_event);
                                            //((ComboBox)ribbon_stacked_items[i]).DropDownClosed += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownClosedEventArgs>(combo_box_event);
                                            //((ComboBox)ribbon_stacked_items[i]).DropDownOpened += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs>(combo_box_event);
                                        }

                                        if (ribbon_stacked_items[i].ItemType.ToString() == "TextBox")
                                        {
                                            if (item_list[i].Attributes["prompttext"].Value != "")
                                            {
                                                ((TextBox)ribbon_stacked_items[i]).PromptText = item_list[i].Attributes["prompttext"].Value;
                                            }

                                            if (item_list[i].Attributes["showimage"].Value == "true")
                                            {
                                                ((TextBox)ribbon_stacked_items[i]).ShowImageAsButton = true;
                                            }

                                            // Use this Event Handler Template for a Text Box

                                            // textbox.EnterPressed += new EventHandler<Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs>(text_box_event);
                                        }
                                    }
                                }
                            }
                        }

                        if (button_structure.Name == "splitbuttons")
                        {
                            XmlNodeList button_list = button_structure.SelectNodes("button");

                            // Make split button object

                            SplitButtonData split_button_data = new SplitButtonData("SplitButton", "Split");
                            SplitButton split_button = ribbon_panel.AddItem(split_button_data) as SplitButton;

                            foreach (XmlNode button in button_list)
                            {
                                split_button.AddPushButton(add_button(button, this_assembly_path));
                            }
                        }

                        if (button_structure.Name == "slideoutpanel")
                        {
                            XmlNodeList button_list = button_structure.SelectNodes("button");

                            // Make slideout panel

                            ribbon_panel.AddSlideOut();

                            foreach (XmlNode button in button_list)
                            {
                                ribbon_panel.AddItem(add_button(button, this_assembly_path));
                            }
                        }

                        if (button_structure.Name == "radiobuttons")
                        {
                            // The following button fields are required

                            if (button_structure.Attributes["name"].Value == "")
                            {
                                TaskDialog button_error = new TaskDialog("Required field for a radio button is missing");
                                button_error.Show();
                            }

                            RadioButtonGroupData radio_button_data = new RadioButtonGroupData(button_structure.Attributes["name"].Value);
                            RadioButtonGroup radio_button_group = ribbon_panel.AddItem(radio_button_data) as RadioButtonGroup;

                            XmlNodeList toggle_button_list = button_structure.ChildNodes;

                            foreach (XmlNode toggle_button in toggle_button_list)
                            {
                                // The following button fields are required

                                if (toggle_button.Attributes["name"].Value == "" || toggle_button.Attributes["text"].Value == "")
                                {
                                    TaskDialog button_error = new TaskDialog("Required fields for a toggle button are missing");
                                    button_error.Show();
                                }
                                string toggle_button_name_string = toggle_button.Attributes["name"].Value;
                                string toggle_button_text_string = toggle_button.Attributes["text"].Value;
                                ToggleButtonData toggle_button_data = new ToggleButtonData(toggle_button_name_string, toggle_button_text_string);

                                // The following button fields are optional
                                // If they aren't empty, do the following

                                if (toggle_button.Attributes["tooltip"].Value != "")
                                {
                                    string toggle_button_tooltip_string = toggle_button.Attributes["tooltip"].Value;
                                    toggle_button_data.ToolTip = toggle_button_tooltip_string;
                                }

                                if (toggle_button.Attributes["largeimage"].Value != "")
                                {
                                    string toggle_button_largeimage_string = toggle_button.Attributes["largeimage"].Value;
                                    toggle_button_data.LargeImage = new BitmapImage(new Uri(toggle_button_largeimage_string));
                                }

                                radio_button_group.AddItem(toggle_button_data);
                            }
                        }
                    }
                }
            }
        }

        static TextBoxData add_textbox(XmlNode textbox_node)
        {
            // The following button fields are required

            if (textbox_node.Attributes["name"].Value == "")
            {
                TaskDialog button_error = new TaskDialog("Required field for a Text Box is missing");
                button_error.Show();
            }

            string textbox_name_string = textbox_node.Attributes["name"].Value;
            TextBoxData textbox_data = new TextBoxData(textbox_name_string);

            // The following button fields are optional
            // If they aren't empty, do the following

            if (textbox_node.Attributes["longdescription"].Value != "")
            {
                string textbox_longdescription_string = textbox_node.Attributes["longdescription"].Value;
                textbox_data.LongDescription = textbox_longdescription_string;
            }

            if (textbox_node.Attributes["image"].Value != "")
            {
                string textbox_image_string = textbox_node.Attributes["image"].Value;
                textbox_data.Image = new BitmapImage(new Uri(textbox_image_string));
            }

            if (textbox_node.Attributes["tooltip"].Value != "")
            {
                string textbox_tooltip_string = textbox_node.Attributes["tooltip"].Value;
                textbox_data.ToolTip = textbox_tooltip_string;
            }

            if (textbox_node.Attributes["tooltipimage"].Value != "")
            {
                string textbox_tooltipimage_string = textbox_node.Attributes["tooltipimage"].Value;
                textbox_data.ToolTipImage = new BitmapImage(new Uri(textbox_tooltipimage_string));
            }

            return textbox_data;
        }

        static void add_combobox(XmlNode combobox_node, ComboBox combobox_item)
        {
            if (combobox_node.Attributes["itemtext"].Value != "")
            {
                string combobox_itemtext_string = combobox_node.Attributes["itemtext"].Value;
                combobox_item.ItemText = combobox_itemtext_string;
            }

            if (combobox_node.Attributes["longdescription"].Value != "")
            {
                string combobox_longdescription_string = combobox_node.Attributes["longdescription"].Value;
                combobox_item.LongDescription = combobox_longdescription_string;
            }

            if (combobox_node.Attributes["tooltip"].Value != "")
            {
                string combobox_tooltip_string = combobox_node.Attributes["tooltip"].Value;
                combobox_item.ToolTip = combobox_tooltip_string;
            }

            if (combobox_node.Attributes["tooltipimage"].Value != "")
            {
                string combobox_tooltip_image_string = combobox_node.Attributes["tooltipimage"].Value;
                combobox_item.ToolTipImage = new BitmapImage(new Uri(combobox_tooltip_image_string));

            }

            XmlNodeList combobox_member_list = combobox_node.ChildNodes;

            foreach (XmlNode combobox_member in combobox_member_list)
            {
                // The following button fields are required

                if (combobox_member.Attributes["name"].Value == "" || combobox_member.Attributes["text"].Value == "")
                {
                    TaskDialog button_error = new TaskDialog("Required fields for a Combo Box Member are missing");
                    button_error.Show();
                }

                string combobox_member_name_string = combobox_member.Attributes["name"].Value;
                string combobox_member_text_string = combobox_member.Attributes["text"].Value;
                ComboBoxMemberData combobox_member_item = new ComboBoxMemberData(combobox_member_name_string, combobox_member_text_string);

                // The following button fields are optional
                // If they aren't empty, do the following

                if (combobox_member.Attributes["image"].Value != "")
                {
                    string combobox_member_image_string = combobox_member.Attributes["image"].Value;
                    combobox_member_item.Image = new BitmapImage(new Uri(combobox_member_image_string));
                }

                if (combobox_member.Attributes["groupname"].Value != "")
                {
                    string combobox_member_groupname_string = combobox_member.Attributes["groupname"].Value;
                    combobox_member_item.GroupName = combobox_member_groupname_string;
                }

                // Add combo box member to combo box

                combobox_item.AddItem(combobox_member_item);
            }

        }
        static PulldownButtonData add_pulldown(XmlNode pulldown_button_node)
        {
            // The following button fields are required

            if (pulldown_button_node.Attributes["name"].Value == "" || pulldown_button_node.Attributes["image"].Value == "")
            {
                TaskDialog button_error = new TaskDialog("Required fields for a Pulldown Button are missing");
                button_error.Show();
            }

            string pulldown_button_name_string = pulldown_button_node.Attributes["name"].Value;
            string pulldown_button_image_string = pulldown_button_node.Attributes["image"].Value;

            // Make pulldown button object

            PulldownButtonData pulldown_button_data = new PulldownButtonData("PullDown", pulldown_button_name_string);
            pulldown_button_data.LargeImage = new BitmapImage(new Uri(pulldown_button_image_string));

            return pulldown_button_data;

        }

        static PushButtonData add_button(XmlNode button, string assembly_path)
        {
            // The following button fields are required

            if (button.Attributes["name"].Value == "" || button.Attributes["classname"].Value == "" || button.Attributes["text"].Value == "")
            {
                TaskDialog button_error = new TaskDialog("Required fields for a button are missing");
                button_error.Show();
            }

            string button_name_string = button.Attributes["name"].Value;
            string button_classname_string = button.Attributes["classname"].Value;
            string button_text_string = button.Attributes["text"].Value;

            PushButtonData push_button_data = new PushButtonData(button_name_string, button_text_string, assembly_path, button_classname_string);

            // The following button fields are optional
            // If they aren't empty, do the following 

            if (button.Attributes["tooltip"].Value != "")
            {
                string button_tooltip_string = button.Attributes["tooltip"].Value;
                push_button_data.ToolTip = button_tooltip_string;
            }

            if (button.Attributes["image"].Value != "")
            {
                string button_image_string = button.Attributes["image"].Value;
                push_button_data.Image = new BitmapImage(new Uri(button_image_string));
            }

            if (button.Attributes["largeimage"].Value != "")
            {
                string button_large_image_string = button.Attributes["largeimage"].Value;
                push_button_data.LargeImage = new BitmapImage(new Uri(button_large_image_string));
            }

            if (button.Attributes["tooltipimage"].Value != "")
            {
                string button_tooltip_image_string = button.Attributes["tooltipimage"].Value;
                push_button_data.ToolTipImage = new BitmapImage(new Uri(button_tooltip_image_string));
            }

            if (button.Attributes["contexthelp"].Value != "")
            {
                string button_context_help_string = button.Attributes["contexthelp"].Value;
                push_button_data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, button_context_help_string));
            }

            return push_button_data;
        }

        // Use this Event Handler Template for a Text Box

        //static void text_box_event(object sender, Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs args)
        //{
        // // write code here
        //}

        // Use this Event Handler Template for a Combo Box

        //static void combo_box_event(object sender, Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs args)
        //{
        //    // write code here
        //}

        //static void combo_box_event(object sender, Autodesk.Revit.UI.Events.ComboBoxDropDownClosedEventArgs args)
        //{
        //    // write code here
        //}

        //static void combo_box_event(object sender, Autodesk.Revit.UI.Events.ComboBoxDropDownOpenedEventArgs args)
        //{
        //    // write code here
        //}
    }
}