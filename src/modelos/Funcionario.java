package modelos;

public class Funcionario {
    private int id;
    private String paterno;
    private String materno;
    private String nombres;
    private String fechaNac;
    private int ci;
    private String genero;

    public Funcionario(int id, String paterno, String materno, String nombres, String fechaNac, int ci, String genero) {
        this.id = id;
        this.paterno = paterno;
        this.materno = materno;
        this.nombres = nombres;
        this.fechaNac = fechaNac;
        this.ci = ci;
        this.genero = genero;
    }

    public int getId() { return id; }
    public String getNombres() { return nombres; }
    public String getPaterno() { return paterno; }
    public String getMaterno() { return materno; }
    public String getFechaNac() { return fechaNac; }
    public int getCi() { return ci; }
    public String getGenero() { return genero; }
}
