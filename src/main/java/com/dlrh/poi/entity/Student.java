package com.dlrh.poi.entity;

public class Student {
    private String id;
    private String name;
    private String address;
    private String gender;
    private String number;
    private Integer age;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Student(String id, String name, String address, String gender, String number, Integer age) {
        this.id = id;
        this.name = name;
        this.address = address;
        this.gender = gender;
        this.number = number;
        this.age = age;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", address='" + address + '\'' +
                ", gender='" + gender + '\'' +
                ", number='" + number + '\'' +
                ", age=" + age +
                '}';
    }
}
