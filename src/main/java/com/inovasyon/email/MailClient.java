package com.inovasyon.email;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;

public class MailClient {

	public static void main(String[] args) throws IOException {

		ExchangeService service = null;

		try {
			service = new ExchangeService(ExchangeVersion.Exchange2010);
			service.setUrl(new URI("https://mail.havelsan.com.tr/ews/exchange.asmx"));
		} catch (Exception e) {

			e.printStackTrace();

		}

		List<String> list = readExcelFile();
		Credentials.setExchangeCredentials(service);
		MailClient.sendEmails(list, service);

	}

	public static List<String> readExcelFile() throws IOException {

		List<String> list = new ArrayList<>();

		FileInputStream fis = new FileInputStream("C:\\Users\\omer\\Desktop\\testMailList.xls");

		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		HSSFSheet spreadsheet = workbook.getSheetAt(0);

		for (Row cells : spreadsheet) {
			HSSFRow row = (HSSFRow) cells;
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				System.out.println(cell.getStringCellValue());
				list.add(cell.getStringCellValue());
			}
			System.out.println();
		}
		fis.close();

		return list;
	}

	public static void sendEmails(List<String> recipientsList, ExchangeService service) {
		try {
			StringBuilder strBldr = new StringBuilder();
			strBldr.append("cello-mail-client-test:");
			strBldr.append(Calendar.getInstance().getTime().toString() + " .");
			strBldr.append("Thanks and Regards");
			strBldr.append("cello-mail-client-test");
			EmailMessage message = new EmailMessage(service);
			message.setSubject("Test sending email");

			message.setBody(new MessageBody(strBldr.toString()));
			for (String string : recipientsList) {
				message.getToRecipients().add(string);
			}
			message.sendAndSaveCopy();
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("message sent");
	}

}