using Microsoft.Office.Interop.Access;
using System.Windows.Automation;

namespace Northwind20.UI.Tests.Components
{
    public static class AutmoationElementExtensions
    {
        public static IEnumerable<AutomationElement> GetChildElementsByName(this AutomationElement parent, string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return parent.FindAll(TreeScope.Children, propCondition).Cast<AutomationElement>();
        }

        public static AutomationElement? GetFirstDescendentElementByName(this AutomationElement parent, string name)
        {
            var results = parent.FindAll(TreeScope.Children, Condition.TrueCondition).Cast<AutomationElement>();

            var stack = new Stack<AutomationElement>();
            stack.Push(parent);
            while (stack.Count > 0)
            {
                var current = stack.Pop();
                if (current.Current.Name == name) { return current; }

                foreach (var child in current.FindAll(TreeScope.Children, Condition.TrueCondition).Cast<AutomationElement>())
                {
                    stack.Push(child);
                }
            }

            return null;
        }

        public static AutomationElement? GetFirstDescendentElementByNameAndType(this AutomationElement parent, string name, ControlType controlType)
        {
            var results = parent.FindAll(TreeScope.Children, Condition.TrueCondition).Cast<AutomationElement>();

            var stack = new Stack<AutomationElement>();
            stack.Push(parent);
            while (stack.Count > 0)
            {
                var current = stack.Pop();
                if (current.Current.Name == name && current.Current.ControlType == controlType) { return current; }

                foreach (var child in current.FindAll(TreeScope.Children, Condition.TrueCondition).Cast<AutomationElement>())
                {
                    stack.Push(child);
                }
            }

            return null;
        }

        public static IEnumerable<AutomationElement> GetChildElementsByControlType(this AutomationElement parent, ControlType controlType)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty, controlType);

            return parent.FindAll(TreeScope.Children, propCondition).Cast<AutomationElement>();
        }


        public static void ChangeTextValue(this AutomationElement element, string value)
        {
            // Set focus for input functionality and begin.
            if (element.Current.ControlType == ControlType.ComboBox)
            {
                element = element.GetChildElementsByControlType(ControlType.Edit).First();
            }

            // A series of basic checks prior to attempting an insertion.
            //
            // Check #1: Is control enabled?
            // An alternative to testing for static or read-only controls
            // is to filter using
            // PropertyCondition(AutomationElement.IsEnabledProperty, true)
            // and exclude all read-only text controls from the collection.
            if (!element.Current.IsEnabled)
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + " is not enabled.\n\n");
            }

            // Check #2: Are there styles that prohibit us
            //           from sending text to this control?
            if (!(element.Current.IsKeyboardFocusable || element.Current.HasKeyboardFocus))
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + "is read-only.\n\n");
            }


            // Control supports the ValuePattern pattern so we can
            // use the SetValue method to insert content.
            else
            {
                if (element.Current.HasKeyboardFocus)
                {

                    SendKeys.SendWait(value);
                }
                else
                {
                    element.SetFocus();

                    if (!element.TryGetCurrentPattern(
                        ValuePattern.Pattern, out var valuePattern))
                    {
                        throw new InvalidOperationException(
                            "The control with an AutomationID of "
                            + element.Current.AutomationId.ToString()
                            + "does not support value pattern.\n\n");
                    }

                    ((ValuePattern)valuePattern).SetValue(value);
                }
            }
        }

        public static void SelectValue(this AutomationElement element, string value)
        {

            // A series of basic checks prior to attempting an insertion.
            //
            // Check #1: Is control enabled?
            // An alternative to testing for static or read-only controls
            // is to filter using
            // PropertyCondition(AutomationElement.IsEnabledProperty, true)
            // and exclude all read-only text controls from the collection.
            if (!element.Current.IsEnabled)
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + " is not enabled.\n\n");
            }

            // Check #2: Are there styles that prohibit us
            //           from sending text to this control?
            if (!(element.Current.IsKeyboardFocusable || element.Current.HasKeyboardFocus))
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + "is read-only.\n\n");
            }

            // Control does not support the ValuePattern pattern
            // so use keyboard input to insert content.
            //
            //
            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expandPattern))
            {
                ((ExpandCollapsePattern)expandPattern).Expand();

                AutomationElement listItem = element.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, value));
                if(listItem.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectPattern))
                {
                    ((SelectionItemPattern)selectPattern).Select();
                }

                ((ExpandCollapsePattern)expandPattern).Collapse();
            }
        }

        public static bool ChildElementExistsByName(this AutomationElement parent, string name)
        {
            Condition propCondition = new PropertyCondition(
                AutomationElement.NameProperty, name, PropertyConditionFlags.IgnoreCase);

            return parent.FindFirst(TreeScope.Children, propCondition) != null;
        }

        public static bool ElementIsKeyboardFocusableCombobox(this AutomationElement element)
        {
            if (element.Current.ControlType != ControlType.ComboBox) return false;
            return element.GetChildElementsByControlType(ControlType.Edit).FirstOrDefault()?.Current.IsKeyboardFocusable ?? false;
        }

        public static void InvokeClick(this AutomationElement element)
        {
            element.TryGetCurrentPattern(InvokePattern.Pattern, out var objPattern);
            InvokePattern invPattern = objPattern as InvokePattern;
            if (invPattern != null)
            {
                invPattern.Invoke();
            }
        }

        public static WindowPattern GetWindowPattern(this AutomationElement targetControl)
        {
            WindowPattern windowPattern = null;

            try
            {
                windowPattern = targetControl.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
            }
            catch (InvalidOperationException)
            {
                // object doesn't support the WindowPattern control pattern
                return null;
            }
            // Make sure the element is usable.
            if (false == windowPattern.WaitForInputIdle(10000))
            {
                // Object not responding in a timely manner
                return null;
            }
            return windowPattern;
        }
    }
}