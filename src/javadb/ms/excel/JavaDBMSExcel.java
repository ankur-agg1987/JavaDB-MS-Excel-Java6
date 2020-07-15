package javadb.ms.excel;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

/**
 *
 * @author Ankur
 */
public class JavaDBMSExcel {

    public static Connection conn = null;
    public static Statement stmt = null;

    public static void main(String[] args) throws SQLException {

        try {
            conn = getConnection();
            stmt = conn.createStatement();
            System.out.println("statement created");
            String excelQuery = "insert into [Sheet1$](FirstName, LastName) values('A', 'K')";
            stmt.executeUpdate(excelQuery);
            System.out.println("query executed");
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (Exception e1) {
            e1.printStackTrace();
        } finally {
        }
    }

    public static Connection getConnection() throws Exception {
        System.out.println("getconnection function called");
        String driver = "sun.jdbc.odbc.JdbcOdbcDriver";
        String url = "jdbc:odbc:excelDB";
        String username = "";
        String password = "";
        Class.forName(driver); // load JDBC-ODBC driver
        System.out.println("driver registered");
        return DriverManager.getConnection(url, username, password);
    }
}
