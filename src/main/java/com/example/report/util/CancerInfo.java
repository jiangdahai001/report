package com.example.report.util;

public class CancerInfo {
  private String name;
  private Boolean value;
  private String apple;
  private String orange;
  private String banana;

  public CancerInfo(String name, Boolean value, String apple, String orange, String banana) {
    this.name = name;
    this.value = value;
    this.apple = apple;
    this.orange = orange;
    this.banana = banana;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public Boolean getValue() {
    return value;
  }

  public void setValue(Boolean value) {
    this.value = value;
  }

  public String getApple() {
    return apple;
  }

  public void setApple(String apple) {
    this.apple = apple;
  }

  public String getOrange() {
    return orange;
  }

  public void setOrange(String orange) {
    this.orange = orange;
  }

  public String getBanana() {
    return banana;
  }

  public void setBanana(String banana) {
    this.banana = banana;
  }
}
