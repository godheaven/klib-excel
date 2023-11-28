package cl.kanopus.excel.writer;

public class ProductTO {

    private final String code;
    private final String name;
    private final double price;

    public ProductTO(String code, String name, double price) {
        this.code = code;
        this.name = name;
        this.price = price;
    }

    public String getCode() {
        return code;
    }

    public String getName() {
        return name;
    }

    public double getPrice() {
        return price;
    }

}
