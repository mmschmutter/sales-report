/**
 * Call object
 *
 * @author Mordechai Schmutter
 * @version 1.0
 */

class Product implements Comparable<Product> {
    private String name;
    private int month;
    private int year;
    private int quantity;

    Product(String name, int month, int year, int quantity) {
        this.name = name;
        this.month = month;
        this.year = year;
        this.quantity = quantity;
    }

    String getName() {
        return name;
    }

    int getMonth() {
        return month;
    }

    int getYear() {
        return year;
    }

    int getQuantity() {
        return quantity;
    }

    String getID() {
        return name + "-" + month + "-" + year;
    }

    void increaseQuantity(int amount) {
        quantity += amount;
    }

    // use quantity to sort Product objects
    public int compareTo(Product other) {
        if (quantity > other.getQuantity()) {
            return -1;
        }
        if (quantity < other.getQuantity()) {
            return 1;
        }
        return 0;
    }
}