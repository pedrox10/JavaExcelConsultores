package modelos;

public class Contrato {
    private int id;
    private String minuta;

    public Contrato(int id, String minuta) {
        this.id = id;
        this.minuta = minuta;
    }

    public int getId() { return id; }
    public String getMinuta() { return minuta; }
}
