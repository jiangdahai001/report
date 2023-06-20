package com.example.report.util;

public class PatientInfo {
  private String name;
  private String phone;
  private String gender;
  private int age;
  public PatientInfo() {}
  public PatientInfo(String name, String phone, String gender, int age) {
    this.name = name;
    this.phone = phone;
    this.gender = gender;
    this.age = age;
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
}
