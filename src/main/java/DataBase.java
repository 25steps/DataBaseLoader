import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import java.sql.*;
import java.util.ArrayList;

public class DataBase {
    private static Connection connection;
    private static Statement statement;

    public static void connection(String path) throws SQLException {
        try {
            Class.forName("org.sqlite.JDBC");
            connection = DriverManager.getConnection("jdbc:sqlite:" + path);
            statement = connection.createStatement();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }

    }

    public static ArrayList getNameTables() throws SQLException {
        ArrayList array = new ArrayList();
        String query = String.format("select name From sqlite_master;");
        ResultSet rs = statement.executeQuery(query);
        while (rs.next())
            if (!rs.getString(1).startsWith("sqlite_") && !rs.getString(1).startsWith("IFK_"))
                array.add(rs.getString(1));
        return array;
    }

    public static ArrayList getColNameTable(String nameTable) throws SQLException {
        ArrayList array = new ArrayList();
        String query = String.format("select * from " + nameTable + " limit 0;");
        ResultSet rs = statement.executeQuery(query);
        ResultSetMetaData data = rs.getMetaData();
        for (int i = 1; i <= data.getColumnCount() ; i++) {
            array.add(data.getColumnName(i));
        }
        return array;
    }

    public static ObservableList getFullTable(String nameTable) throws SQLException {
        ObservableList data = FXCollections.observableArrayList();
        String query = String.format("select * from " + nameTable + ";");
        ResultSet rs = statement.executeQuery(query);
        while (rs.next()){
            ObservableList<String> row = FXCollections.observableArrayList();
            for (int i = 1; i <= rs.getMetaData().getColumnCount() ; i++) {
                if (rs.getString(i) == null)
                    row.add("");
                else
                    row.add(rs.getString(i));
            }
            data.add(row);
        }
        return data;
    }


    public static void disconnect(){
        if (connection != null) {
            try {
                statement.close();
                connection.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

}
