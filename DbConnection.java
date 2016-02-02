package maid.myapi.com;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class DbConnection {
 
    public static Connection getConnection()
    {
     Connection con = null;
     
     try 
     {
    Class.forName("com.mysql.jdbc.Driver");
    con  =DriverManager.getConnection("jdbc:mysql://localhost:3306/MaidDB","root","root");
    System.out.println("Database connection established.");
     } 
     catch (ClassNotFoundException e) 
     {
   // TODO Auto-generated catch block
      e.printStackTrace();
     } 
     catch (SQLException e) 
     {
   // TODO Auto-generated catch block
      e.printStackTrace();
     }
     
     return con; 
     
    }
 
     public static void closeConnection(Connection con){      
      if(con!=null){
       try{
       con.close();
       }catch(SQLException e){}
      }      
     }   
       public static void closePreparedStatement(PreparedStatement ps){      
      if(ps!=null){
       try{
       ps.close();
       }catch(SQLException e){}
      }
      
     }

}
