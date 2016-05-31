package encuestamagjava;

import BD.Conexion;
import java.util.List;

/**
 *
 * @author ArmanJB
 */
public class EncuestaMAGJAVA {
    private Conexion  bd;
    
    public void conectar(){
        bd = new Conexion();   
    }
    
    public boolean NuevoRegistro(List<String> campos){
        bd.conecta();
        return bd.crearRegistro(campos);
    }
    
}
