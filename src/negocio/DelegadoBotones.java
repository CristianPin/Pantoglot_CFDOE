package negocio;

import datos.ControlArchivos;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import presentacion.FormGestionArchivos;

public class DelegadoBotones extends Thread {

    public void ExaminarOrigen(FormGestionArchivos aThis) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int resultado = chooser.showOpenDialog(aThis);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            FormGestionArchivos.TOrigen.setText(chooser.getSelectedFile().toString());
        }
    }

    public void ExaminarDestino(FormGestionArchivos aThis) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int resultado = chooser.showOpenDialog(aThis);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            FormGestionArchivos.TDestino.setText(chooser.getSelectedFile().toString());
        }
    }

    public void Excel(FormGestionArchivos aThis) {
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx");
        chooser.setFileFilter(filter);
        //chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        int resultado = chooser.showOpenDialog(aThis);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            FormGestionArchivos.TExcel.setText(chooser.getSelectedFile().toString());
        }
    }


    @Override
    public void run() {
        ControlArchivos Control = new ControlArchivos();
        File FinalRuta = new File(FormGestionArchivos.TOrigen.getText());
        Control.NewRuta = new File(FormGestionArchivos.TDestino.getText());
        ControlArchivos.nombreArchivo = FormGestionArchivos.TExcel.getText();
        
        Control.LeerEXCEL();
        Control.Carpetas(FinalRuta, 1);
        
        FormGestionArchivos.BEjecutar.setEnabled(true);
        FormGestionArchivos.PBEjecutar.setIndeterminate(false);
        FormGestionArchivos.PBEjecutar.setValue(100);
    }
    
    

}
