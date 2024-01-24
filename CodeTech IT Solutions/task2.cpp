// Task 2 : Simple Calculator

#include<bits/stdc++.h>
using namespace std;

int main()
{
    int ch;
    float no1,no2;

    //Input values for Operation
    cout<<"Enter two numbers for Calculation: \n";
    cin>>no1>>no2;
    cout<<"Choose the Operation: \n";
    cout<<"1. Addition \n2. Subtraction \n3. Multiplication \n4. Division\n";
    cin>>ch;

    //Perform the operation
    switch(ch){
        case 1:
            cout<<"Addition: "<<no1<<"+"<<no2 <<"=" <<no1+no2;      //Result of Addition
            break;
        case 2:
            cout<<"Subtraction: "<<no1<<"-"<<no2<<"=" << no1-no2;   //Result of Subtraction
            break;
        case 3:
            cout<<"Multiplication: "<<no1<<"*"<<no2<<"="<< no1*no2;     //Result of Multiplication
            break;
        case 4:
            cout<<"Division: "<<no1<<"/"<<no2<<"="<< no1/no2;       //Result of Division
            break;
        default:
            cout<<"Wrong Choice, Incorrect Operation";
    }
    
    return 0;
}