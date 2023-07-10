package com.example.report.util;

import java.util.List;

public class PatientInfo {
  private String name;
  private String phone;
  private String gender;
  private int age;
  private List<String> geneList;
  private List<String> picList;
  public PatientInfo() {}
  public PatientInfo(String name, String phone, String gender, int age, List<String> geneList, List<String> picList) {
    this.name = name;
    this.phone = phone;
    this.gender = gender;
    this.age = age;
    this.geneList = geneList;
    this.picList = picList;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public String getPhone() {
    return phone;
  }

  public void setPhone(String phone) {
    this.phone = phone;
  }

  public String getGender() {
    return gender;
  }

  public void setGender(String gender) {
    this.gender = gender;
  }

  public int getAge() {
    return age;
  }

  public void setAge(int age) {
    this.age = age;
  }

  public List<String> getGeneList() {
    return geneList;
  }

  public void setGeneList(List<String> geneList) {
    this.geneList = geneList;
  }

  public List<String> getPicList() {
    return picList;
  }

  public void setPicList(List<String> picList) {
    this.picList = picList;
  }
}
