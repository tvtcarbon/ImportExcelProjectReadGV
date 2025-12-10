package de.srs.importExcelMRP.obj;

public class MrpRezObjExt extends MrpRezObj{
	private String einsatzMaterial;
	private Float mi;
	private Float moto1_2;
	private Float ip;
	private Float matKosten;

	public String getEinsatzMaterial() {
		return einsatzMaterial;
	}

	public void setEinsatzMaterial(String einsatzMaterial) {
		this.einsatzMaterial = einsatzMaterial;
	}
	
	public Float getMi() {
		return mi;
	}

	public Float getMoto1_2() {
		return moto1_2;
	}

	public void setMoto1_2(Float moto1) {
		this.moto1_2 = moto1;
	}

	public void setMi(Float mi) {
		this.mi = mi;
	}
	
	public Float getMatKosten() {
		return matKosten;
	}

	public void setMatKosten(Float matKosten) {
		this.matKosten = matKosten;
	}

	public Float getIp() {
		return ip;
	}

	public void setIp(Float ip) {
		this.ip = ip;
	}

	public MrpRezObjExt() {
		super();
	}

	@Override
	public String toString() {
		return "MrpRezObjExt [einsatzMaterial=" + einsatzMaterial + ", mi=" + mi + ", moto1_2=" + moto1_2 + ", ip=" + ip
				+ ", matKosten=" + matKosten + "]";
	}


	

}
