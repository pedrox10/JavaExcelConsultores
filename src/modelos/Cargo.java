package modelos;

public class Cargo {
    private int id;
    private String nombre;
    public int nivel;
    public int idPartida;

    public Cargo(int id, String nombre, int nivel, int id_partida) {
        this.id = id;
        this.nombre = nombre;
        this.nivel = nivel;
        this.idPartida = id_partida;
    }

    public int getId() { return id; }
    public String getNombre() { return nombre; }
    public int getNivel() { return nivel; }
    public int getIdPartida() { return idPartida; }
}
