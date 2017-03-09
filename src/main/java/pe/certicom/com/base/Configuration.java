package pe.certicom.com.base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class Configuration {

    Properties props;

    public Configuration() {
        try {
            props = new Properties();
            props.load(new FileInputStream(Constantes.FILE));
        } catch (FileNotFoundException e) {
            System.out.println("Error, El archivo no exite");
        } catch (IOException e) {
            System.out.println("Error, No se puede leer el archivo");
        }
    }

    public Properties getProperties() {
        return this.props;
    }
}
