package com.example.report.util;

public class CancerInfo {
  private String name;
  private Boolean value;

  public CancerInfo(String name, Boolean value) {
    this.name = name;
    this.value = value;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public Boolean isValue() {
    return value;
  }

  public void setValue(Boolean value) {
    this.value = value;
  }
}
