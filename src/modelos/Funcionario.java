package modelos;

public class Funcionario {
    private int id;
    private String nombre;
    private String paterno;
    private String materno;
    private String fechaNac;
    private int ci;

    public Funcionario(int id, String nombre, String paterno, String materno, String fechaNac, int ci) {
        this.id = id;
        this.nombre = nombre;
        this.paterno = paterno;
        this.materno = materno;
        this.fechaNac = fechaNac;
        this.ci = ci;
    }

    public int getId() { return id; }
    public String getNombre() { return nombre; }
    public String getPaterno() { return paterno; }
    public String getMaterno() { return materno; }
    public String getFechaNac() { return fechaNac; }
    public int getCi() { return ci; }
}
