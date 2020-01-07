package busy_extn;

import java.util.Date;
import java.util.List;

public class InvoiceData {
	String invoiceNumber;
	String gstNumber;
	String receiverName;
	Date invoiceDate;
	String placeOfSupply;
	String reverseCharge;
	String invoiceType;
	List<TaxData> taxes;
}
