package modelos;

public class PartidaSalud {
    private int id;
    private String nombre;
    private String fuente;
    private String organismo;
    private int categoria;

    public PartidaSalud(int id, String nombre, String fuente, String organismo, int categoria) {
        this.id = id;
        this.nombre = nombre;
        this.fuente = fuente;
        this.organismo = organismo;
        this.categoria = categoria;
    }

    public int getId() { return id; }
    public String getNombre() { return nombre; }

    public String getFuente() {
        return fuente;
    }

    public String getOrganismo() {
        return organismo;
    }

    public int getCategoria() {
        return categoria;
    }
}