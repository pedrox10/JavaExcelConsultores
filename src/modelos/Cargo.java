package modelos;

public class Cargo {
    private int id;
    private String nombre;

    public Cargo(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

    public int getId() { return id; }
    public String getNombre() { return nombre; }
}
