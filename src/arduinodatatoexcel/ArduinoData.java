package arduinodatatoexcel;

/**
 * @version 1.0
 * @author Seijas
 */
public class ArduinoData {
    
    private final String temp, hum, state;
    
    public ArduinoData(String data){
        String vec[] = data.split(";");
        hum = vec[0];
        temp = vec[1];
        state = vec[2];
    }

    public String getTemp() {
        return temp;
    }

    public String getHum() {
        return hum;
    }

    public String getState() {
        return state;
    }
    
}
