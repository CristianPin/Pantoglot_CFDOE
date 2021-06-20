package negocio;

import presentacion.FormGestionArchivos;

public class BarraProgreso extends Thread{

    @Override
    public void run() {
        FormGestionArchivos.BEjecutar.setEnabled(false);
        FormGestionArchivos.PBEjecutar.setIndeterminate(true);
    }
    
    
    
}
