package modelos;

public class Contrato {
    private int id;
    private String minuta;
    private String fechaInicio;
    private String fechaConclusion;
    private int monto;

    public Contrato(int id, String minuta, String fechaInicio, String fechaConclusion, int monto) {
        this.id = id;
        this.minuta = minuta;
        this.fechaInicio = fechaInicio;
        this.fechaConclusion = fechaConclusion;
        this.monto = monto;
    }

    public int getId() { return id; }
    public String getMinuta() { return minuta; }
    public String getFechaInicio() { return fechaInicio; }
    public String getFechaConclusion() { return fechaConclusion; }
    public int getMonto() { return monto; }
}
