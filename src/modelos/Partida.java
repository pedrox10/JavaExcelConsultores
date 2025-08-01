package modelos;

public class Partida {
    private int id;
    private String nombre;

    public Partida(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

    public int getId() { return id; }
    public String getNombre() { return nombre; }
}
