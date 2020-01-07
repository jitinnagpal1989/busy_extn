package busy_extn;

public class TaxData {
	double invoiceValue;
	public double getInvoiceValue() {
		return invoiceValue;
	}
	public void setInvoiceValue(double invoiceValue) {
		this.invoiceValue = invoiceValue;
	}
	double rate;
	double taxableVal;
	double totalTax;
	public double getRate() {
		return rate;
	}
	public void setRate(double rate) {
		this.rate = rate;
	}
	public double getTaxableVal() {
		return taxableVal;
	}
	public void setTaxableVal(double taxableVal) {
		this.taxableVal = taxableVal;
	}
	public double getTotalTax() {
		return totalTax;
	}
	public void setTotalTax(double totalTax) {
		this.totalTax = totalTax;
	}

}
