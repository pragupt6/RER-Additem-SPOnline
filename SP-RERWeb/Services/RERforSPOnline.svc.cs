using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace SP_RERWeb.Services
{
    public class RERforSPOnline : IRemoteEventService
    {
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            // On Item Added event, the list item creation executes
            if (properties.EventType == SPRemoteEventType.ItemAdded)
            {
                using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
                {
                    if (clientContext != null)
                    {
                        try
                        {
                            string title = properties.ItemEventProperties.AfterProperties["Title"].ToString();



                            List lstDemoeventReceiver = clientContext.Web.Lists.GetByTitle(properties.ItemEventProperties.ListTitle);
                            ListItem itemDemoventReceiver = lstDemoeventReceiver.GetItemById(properties.ItemEventProperties.ListItemId);

                            itemDemoventReceiver["Description"] = "Description from RER:: "+title;
                             itemDemoventReceiver.Update();
                            clientContext.ExecuteQuery();
                        }
                        catch (Exception ex) { }
                    }
                }
            }
        }

        /// <summary>
        /// Handles events that occur after an action occurs, such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>

    }
}
