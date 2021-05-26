package com.inovasyon.email;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;

public class Credentials {

    private static String userName="user";
    private static String password="password";

    public static void setExchangeCredentials(ExchangeService service) {
        ExchangeCredentials credentials = new WebCredentials(userName, password, "Hvlnet");
        service.setCredentials(credentials);
    }
}
