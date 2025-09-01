using System.Collections;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;

namespace CostsViewer
{
    public static class SelectedItemsBehavior
    {
        public static readonly DependencyProperty SelectedItemsProperty = DependencyProperty.RegisterAttached(
            "SelectedItems",
            typeof(IList),
            typeof(SelectedItemsBehavior),
            new PropertyMetadata(null, OnSelectedItemsChanged));

        public static void SetSelectedItems(DependencyObject element, IList value)
        {
            element.SetValue(SelectedItemsProperty, value);
        }

        public static IList GetSelectedItems(DependencyObject element)
        {
            return (IList)element.GetValue(SelectedItemsProperty);
        }

        private static readonly DependencyProperty IsUpdatingProperty = DependencyProperty.RegisterAttached(
            "IsUpdating",
            typeof(bool),
            typeof(SelectedItemsBehavior),
            new PropertyMetadata(false));

        private static void SetIsUpdating(DependencyObject element, bool value)
        {
            element.SetValue(IsUpdatingProperty, value);
        }

        private static bool GetIsUpdating(DependencyObject element)
        {
            return (bool)element.GetValue(IsUpdatingProperty);
        }

        private static readonly DependencyProperty CollectionProperty = DependencyProperty.RegisterAttached(
            "Collection",
            typeof(INotifyCollectionChanged),
            typeof(SelectedItemsBehavior),
            new PropertyMetadata(null));

        private static readonly DependencyProperty CollectionChangedHandlerProperty = DependencyProperty.RegisterAttached(
            "CollectionChangedHandler",
            typeof(NotifyCollectionChangedEventHandler),
            typeof(SelectedItemsBehavior),
            new PropertyMetadata(null));

        private static void SetCollection(DependencyObject element, INotifyCollectionChanged? value)
        {
            element.SetValue(CollectionProperty, value);
        }

        private static INotifyCollectionChanged? GetCollection(DependencyObject element)
        {
            return (INotifyCollectionChanged?)element.GetValue(CollectionProperty);
        }

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ListBox listBox)
                return;

            listBox.SelectionChanged -= ListBox_SelectionChanged;

            var oldCollectionChanged = GetCollection(listBox);
            var oldHandler = (NotifyCollectionChangedEventHandler?)listBox.GetValue(CollectionChangedHandlerProperty);
            if (oldCollectionChanged != null && oldHandler != null)
            {
                oldCollectionChanged.CollectionChanged -= oldHandler;
                SetCollection(listBox, null);
                listBox.SetValue(CollectionChangedHandlerProperty, null);
            }

            if (e.NewValue is IList newList)
            {
                listBox.SelectionChanged += ListBox_SelectionChanged;
                listBox.Loaded -= ListBox_Loaded;
                listBox.Loaded += ListBox_Loaded;

                if (newList is INotifyCollectionChanged incc)
                {
                    NotifyCollectionChangedEventHandler handler = (s, args) => SyncListBoxFromCollection(listBox);
                    incc.CollectionChanged += handler;
                    SetCollection(listBox, incc);
                    listBox.SetValue(CollectionChangedHandlerProperty, handler);
                }

                SyncListBoxFromCollection(listBox);
            }
        }

        private static void ListBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is ListBox listBox)
            {
                SyncListBoxFromCollection(listBox);
            }
        }

        private static void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ListBox listBox)
                return;

            if (GetIsUpdating(listBox))
                return;

            var target = GetSelectedItems(listBox);
            if (target == null)
                return;

            SetIsUpdating(listBox, true);

            try
            {
                target.Clear();
                foreach (var item in listBox.SelectedItems)
                {
                    target.Add(item);
                }
            }
            finally
            {
                SetIsUpdating(listBox, false);
            }
        }

        private static void SyncListBoxFromCollection(ListBox listBox)
        {
            var source = GetSelectedItems(listBox);
            if (source == null)
                return;

            if (GetIsUpdating(listBox))
                return;

            SetIsUpdating(listBox, true);
            try
            {
                listBox.SelectedItems.Clear();
                foreach (var item in source)
                {
                    listBox.SelectedItems.Add(item);
                }
            }
            finally
            {
                SetIsUpdating(listBox, false);
            }
        }
    }
}


