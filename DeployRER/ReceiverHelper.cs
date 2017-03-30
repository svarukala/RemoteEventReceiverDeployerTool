﻿using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
//using System.ServiceModel;
//using System.ServiceModel.Channels;
using System.Web;
using System.Configuration;

namespace DeployRER
{
    /// <summary>
    /// Helper Class to Create the Remote Event Receiver
    /// </summary>
    public class ReceiverHelper
    {
        /// <summary>
        /// Creates a Remote Event Receiver
        /// </summary>
        /// <param name="receiverName">The name of the remote event receiver</param>
        /// <param name="type"><see cref="Microsoft.SharePoint.Client.EventReceiverType"/></param>
        /// <returns><see cref="Microsoft.SharePoint.Client.EventReceiverDefinitionRemoCreationInformation"/></returns>
        public static EventReceiverDefinitionCreationInformation CreateEventReciever(string receiverName, EventReceiverType type, string receiverUrl)
        {

            EventReceiverDefinitionCreationInformation _rer = new EventReceiverDefinitionCreationInformation();
            _rer.EventType = type;
            _rer.ReceiverName = receiverName;
            _rer.ReceiverClass = "Mylan.SPOLib.RER.SPOLibRERService";
            //"ECM.AutoTaggingWeb.Services.AutoTaggingService";
            _rer.ReceiverUrl = receiverUrl;
                //"https://a5628bc1.ngrok.io/SPOLibRERService.svc";
                //"https://samplererweb.azurewebsites.net/Services/SampleRERService.svc";
            //"http://localhost:1141/Services/SampleRERService.svc";
            //"https://samplererweb.azurewebsites.net/Services/SampleRERService.svc";
            //"https://rerhost.azurewebsites.net/Services/AutoTaggingService.svc";
            _rer.Synchronization = EventReceiverSynchronization.Synchronous;
            return _rer;
        }

        /// <summary>
        /// Checks to see if a Remote Event Receiver Exists on the list
        /// </summary>
        /// <param name="ctx">An Authenticated ClientClient</param>
        /// <param name="list">The List</param>
        /// <param name="eventReceiverName">The name of the receiver</param>
        /// <returns>a bool indicating if the Receiver exists on the list </returns>
        public static bool DoesEventReceiverExistByName(ClientContext ctx, List list, string eventReceiverName )
        {
            bool _doesExist = false;
            ctx.Load(list, lib => lib.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = list.EventReceivers.Where(e => e.ReceiverName == eventReceiverName).FirstOrDefault();
            if (_rer != null) {
                _doesExist = true;
            }

            return _doesExist;
        }
        public static bool DoesEventReceiverExistByName(ClientContext ctx, Web web, string eventReceiverName)
        {
            bool _doesExist = false;
            ctx.Load(web, w => w.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = web.EventReceivers.Where(e => e.ReceiverName == eventReceiverName).FirstOrDefault();
            if (_rer != null)
            {
                _doesExist = true;
            }

            return _doesExist;
        }

        public static bool DoesEventReceiverExistByName(ClientContext ctx, Site site, string eventReceiverName)
        {
            bool _doesExist = false;
            ctx.Load(site, w => w.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = site.EventReceivers.Where(e => e.ReceiverName == eventReceiverName).FirstOrDefault();
            if (_rer != null)
            {
                _doesExist = true;
            }

            return _doesExist;
        }

        /// <summary>
        /// Add a Remote Event Receiver to a List
        /// </summary>
        /// <param name="ctx">An Authenticated ClientContext</param>
        /// <param name="list">The list</param>
        /// <param name="eventReceiverInfo"><see cref="Microsoft.SharePoint.Client.EventReceiverDefinitionCreationInformation"/></param>
        public static void AddEventReceiver(ClientContext ctx, List list, EventReceiverDefinitionCreationInformation eventReceiverInfo)
        {
            if (!DoesEventReceiverExistByName(ctx, list, eventReceiverInfo.ReceiverName))
            {
                list.EventReceivers.Add(eventReceiverInfo);
                ctx.ExecuteQuery();
            }
        }

        public static void AddEventReceiver(ClientContext ctx, Web web, EventReceiverDefinitionCreationInformation eventReceiverInfo)
        {
            if (!DoesEventReceiverExistByName(ctx, web, eventReceiverInfo.ReceiverName))
            {
                web.EventReceivers.Add(eventReceiverInfo);
                ctx.ExecuteQuery();
            }
        }

        public static void AddEventReceiver(ClientContext ctx, Site site, EventReceiverDefinitionCreationInformation eventReceiverInfo)
        {
            if (!DoesEventReceiverExistByName(ctx, site, eventReceiverInfo.ReceiverName))
            {
                site.EventReceivers.Add(eventReceiverInfo);
                ctx.ExecuteQuery();
            }
        }

        /// <summary>
        /// Remove a Remote Event Receiver from a list
        /// </summary>
        /// <param name="ctx">An Authenticated ClientContext</param>
        /// <param name="list">The List</param>
        /// <param name="receiverName">The Remote Event Receiver name</param>
        public static void RemoveEventReceiver(ClientContext ctx, List list, string receiverName)
        {
            ctx.Load(list, lib => lib.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = list.EventReceivers.Where(e => e.ReceiverName == receiverName).FirstOrDefault();
            if(_rer != null)
            {
                _rer.DeleteObject();
                ctx.ExecuteQuery();
            }
        }

        public static void RemoveEventReceiver(ClientContext ctx, Web web, string receiverName)
        {
            ctx.Load(web, lib => lib.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = web.EventReceivers.Where(e => e.ReceiverName == receiverName).FirstOrDefault();
            if (_rer != null)
            {
                _rer.DeleteObject();
                ctx.ExecuteQuery();
            }
        }

        public static void RemoveEventReceiver(ClientContext ctx, Site site, string receiverName)
        {
            ctx.Load(site, lib => lib.EventReceivers);
            ctx.ExecuteQuery();

            var _rer = site.EventReceivers.Where(e => e.ReceiverName == receiverName).FirstOrDefault();
            if (_rer != null)
            {
                _rer.DeleteObject();
                ctx.ExecuteQuery();
            }
        }
    }
}