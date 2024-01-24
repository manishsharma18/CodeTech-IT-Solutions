// Task 3 : TIC-TAC-TOE GAME

#include <bits/stdc++.h>

using namespace std;

// Function to Show the current state of the board
void dispBoard(const vector<vector<char>>& board) {
    cout << "    0   1   2\n";
    cout << "   -----------\n";
    for (int x = 0; x < 3; x++) {
        cout << x << " | ";
        for (int y = 0; y < 3; y++) {
            cout << board[x][y] << " | ";
        }
        cout << "\n   -----------\n";
    }
}

// Function to Make a Move
bool mkMove(vector<vector<char>>& board, char player, int r, int c) {
    // Check if the chosen cell is empty or not
    if (r >= 0 && r < 3 && c >= 0 && c < 3 && board[r][c] == ' ') {
        board[r][c] = player;
        return true;
    }
    return false;
}

// Function to check for a Win
bool chWin(const vector<vector<char>>& board, char player) {
    // Check all the rows, columns, and diagonals for a win
    for (int x = 0; x < 3; x++) {
        if ((board[x][0] == player && board[x][1] == player && board[x][2] == player) ||
            (board[0][x] == player && board[1][x] == player && board[2][x] == player)) {
            return true;
        }
    }

    return (board[0][0] == player && board[1][1] == player && board[2][2] == player) ||
           (board[0][2] == player && board[1][1] == player && board[2][0] == player);
}

//Function to Check for a Draw
bool chDraw(const vector<vector<char>>& board) {
    // Check if the board is full (no empty cells are available)
    for (const auto& r : board) {
        for (char cell : r) {
            if (cell == ' ') {
                return false;  // There is an empty cell, game is not draw
            }
        }
    }
    return true;  // All cells are filled, game is Draw
}

//Function to Switch between players X and O
void switchPlayers(char& currPlayer) {
    // Switch between Player 'X' and 'O'
    currPlayer = (currPlayer == 'X') ? 'O' : 'X';
}

int main() {
    vector<vector<char>> board(3, vector<char>(3, ' '));
    char currPlayer = 'X';
    bool gameCheck = true;
    
    cout << "\t\e[1mTIC-TAC-TOE GAME: \e[0m\n";
    cout << "    -------------------------\n\n";

    do {
        // Display Current state of the Board
        dispBoard(board);

        // Get Current Player Input
        int r, c;
        cout << "Player : " << currPlayer << ", enter your move (row and column): ";
        cin >> r >> c;

        // Make Move
        if (mkMove(board, currPlayer, r, c)) {
            // Check if the Current Player has Won
            if (chWin(board, currPlayer)) {
                dispBoard(board);
                cout << "Congratulation, Player : " << currPlayer << " Win!\n";
                gameCheck = false;
            } else if (chDraw(board)) {
                // Check for a draw
                dispBoard(board);
                cout << "OOPS, It's a Draw!\n";
                gameCheck = false;
            } else {
                // Switch between players X and O, if the game is still ongoing
                switchPlayers(currPlayer);
            }
        } else {
            cout << "OOPS, Invalid move. Try again!\n";
        }

    } while (gameCheck);

    // Ask if players want to play another game
    char plyAgain;
    cout << "Do you want to play again? (Y/N): ";
    cin >> plyAgain;

    if (plyAgain == 'y' || plyAgain == 'Y') {
        main();  // Restart Game
    } else {
        cout << "Thanks for playing Tic-Tac-Toe!\n";
    }

    return 0;
}