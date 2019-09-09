/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package otchet;

/**
 *
 * @author Dry1d
 */
public class DataModel {
    
    private int id;
    private String date;
    private String time;
    private String st;
    private String direction;
    private String fio;
    private String podr;
    private String id_staff;
    
    public DataModel() {
    }
    
    public DataModel(int id, String date, String time, String st, String direction, String fio, String podr, String id_staff) {
        
        this.id = id;
        this.date = date;
        this.time = time;
        this.st = st;
        this.direction = direction;
        this.fio = fio;
        this.podr = podr;
        this.id_staff = id_staff;
    }
    
    public int getId(){
        return id;
    }
    
    public void setId(int id){
        this.id = id;
    }
    
    public String getDate(){
        return date;
    }
    
    public void setDate(String date){
        this.date = date;
    }
    
    public String getTime(){
        return time;
    }
    
    public void setTime(String time){
        this.time = time;
    }
    
    public String getSt(){
        return st;
    }
    
    public void setSt(String st){
        this.st = st;
    }
    
    public String getDirection(){
        return direction;
    }
    
    public void setDirection(String direction){
        this.direction = direction;
    }
    
    public String getFio(){
        return fio;
    }
    
    public void setFio(String fio){
        this.fio = fio;
    }
    
    public String getPodr(){
        return podr;
    }
    
    public void setPodr(String podr){
        this.podr = podr;
    }
    
    public String getId_staff(){
        return id_staff;
    }
    
    public void setId_staff(String id_staff){
        this.id_staff = id_staff;
    }
}
