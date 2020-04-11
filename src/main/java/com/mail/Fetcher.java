package com.mail;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

import java.net.URI;
import java.util.Arrays;

public class Fetcher {
    public static void main(String[] args) {
        try {
            ExchangeService service = new ExchangeService();
            service.setUrl(new URI("https://owa.luxoft.com/EWS/Exchange.asmx"));
            ExchangeCredentials credentials = new WebCredentials("", "", "");
            service.setCredentials(credentials);
            ItemView view = new ItemView(Integer.MAX_VALUE);
            view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
            Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
            FindItemsResults<Item> results = service.findItems(folder.getId(),view);
            service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
