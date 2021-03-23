package com.java.app.model;

public class Country {

    public static final String TRAVEL_FRIENDLY ="can travel";
    public static final String RESTRICTED_TRAVEL ="restricted travel";
    public static final String QUARANTINE ="quarantine";

	private String name;
	private String shortCode;
    private String travelfriendly;
	
	public Country(String n, String c){
		this.name=n;
		this.shortCode=c;
	}
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getShortCode() {
		return shortCode;
	}
	public void setShortCode(String shortCode) {
		this.shortCode = shortCode;
	}
	
	@Override
	public String toString(){
		return name + "::" + shortCode + "::" + travelfriendly;
	}

    public String getTravelfriendly() {
        return travelfriendly;
    }

    public void setTravelfriendly(String travelfriendly) {
        this.travelfriendly = travelfriendly;
    }
	
}
