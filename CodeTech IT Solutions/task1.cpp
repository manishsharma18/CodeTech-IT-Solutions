// Task 1 : Guess the Number Game

#include<bits/stdc++.h>
using namespace std;

int main()
{
    srand((unsigned int) time(NULL));  // For random number generator
    int num= (rand()%100) + 1;   // Generate a random number between 1 and 100
    int guess=0;

    do
    {
        // Input for Player's guess
        cout<<"Enter Guess(1-100): ";
        cin>>guess;

         // Check the guess for Low, High and Win
        if(guess>num)
            cout<<"Too High!!\n";
        else if(guess<num)
            cout<<"Too Low!!\n";
        else    
            cout<<"Congratulation, You Won!!\n";
    } while (guess!=num);
    
    
    return 0;

}