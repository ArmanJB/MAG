package BD;

import java.sql.*;
import java.util.List;
import javax.swing.JOptionPane;

public class Conexion {
    private Connection conexion = null;
        
    public boolean conecta(){
        boolean resp = false;
        try{
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver").newInstance();
        }catch(Exception ex){
            JOptionPane.showMessageDialog(null, "Falló Driver", "Driver", JOptionPane.ERROR_MESSAGE);
        }
        try{
            String contra = JOptionPane.showInputDialog(null, "Contresaña Usuario sa", "MAG", JOptionPane.QUESTION_MESSAGE);
            //conexion = DriverManager.getConnection("jdbc:sqlserver://localhost;databaseName=MAG;integratedSecurity=true;");
            //conexion = DriverManager.getConnection("jdbc:sqlserver://localhost;databaseName=MAG; user=sa; password=ajb");
            conexion = DriverManager.getConnection("jdbc:sqlserver://localhost;databaseName=MAG; user=sa; password="+contra);
            resp = true;
        }catch (SQLException se){
            JOptionPane.showMessageDialog(null, 
                    "Mensaje: "+se.getMessage()+
                    "\nEstado: "+se.getSQLState()+
                    "\nError: "+se.getErrorCode(),"Error",JOptionPane.ERROR_MESSAGE);
        }
        return resp;
    }
    public void desconecta(){
        try{
            conexion.close();
        }catch(SQLException se){
            
        }
    }
    
    private boolean Update(String sql)  {
        try{
            Statement state = conexion.createStatement();
            //System.out.println(sql);
            state.executeUpdate(sql);
            state.close();
            return true;
        }catch(SQLException se){
            System.err.println(se);
            return false;
        }
    }
    
    public boolean crearRegistro(List<String> campos){
        String query = "INSERT INTO Encuesta VALUES('";
        for (int i = 0; i < campos.size()-1; i++) {
            query += campos.get(i) + "', '";
        }
        query += campos.get(campos.size()-1) + "')";
        
        return Update(query);
    }
        
}
