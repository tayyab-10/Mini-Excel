#include <iostream>
#include <vector>
#include <conio.h>
#include <climits>
#include <fstream>
#include <Windows.h>
#include <sstream>
#include <iomanip>

#define ANSI_COLOR_RED "\033[91m"
#define ANSI_COLOR_YELLOW "\033[93m"
#define ANSI_COLOR_RESET "\033[0m"
#define ANSI_COLOR_GREEN "\033[92m"
#define ANSI_COLOR_BLUE "\x1b[34m"
#define COLOR_BLUE 9
using namespace std;

enum DataType // user-defined data type
{
    CHAR_TYPE,
    INT_TYPE,
    FLOAT_TYPE
};

enum cellAlignment
{
    LEFT,
    RIGHT,
    CENTER
};
void SetConsoleColor(int colorCode)
{
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), colorCode);
}

// Header of the Application

void Header()
{
    system("cls");
    SetConsoleColor(COLOR_BLUE);
    cout << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << "                      ##                      ##     @@@@@@@@@@@    ##                     ##    @@@@@@@@@@@                       @@@@@@@@@@@    @@           @@          @@@@@@@@     @@@@@@@@@@@     @@                             " << endl;
    cout << "                      @@  %%              %%  @@         ##         @@ %%                  @@         ##                          ##               %%        %%           ##            ##              ##                         " << endl;
    cout << "                      ##    %%           %%   ##         ##         ##   %%                ##         @@                          @@                @@     @@            @@             @@              @@                        " << endl;
    cout << "                      @@      %%        %%    @@         ##         @@     %%              @@         ##                          ##                  %%  %%            ##              ##              ##                   " << endl;
    cout << "                      ##        %%     %%     ##         ##         ##       %%            ##         @@                          @@                    @@             @@               @@              @@                     " << endl;
    cout << "                      @@          %%  %%      @@         ##         @@         %%          @@         ##                          ##@@@@@@              %% %%          ##               ##@@@@@@        ##                                    " << endl;
    cout << "                      ##            %%        ##         ##         ##           %%        ##         @@                          @@                   %%   %%         @@               @@              @@                     " << endl;
    cout << "                      @@                      @@         ##         @@             %%      @@         ##                          ##                  @@      @@        ##              ##              ##                  " << endl;
    cout << "                      ##                      ##         ##         ##               %%    ##         @@                          @@                %%         %%        @@             @@              @@                  " << endl;
    cout << "                      @@                      @@         ##         @@                 %%  @@         ##                          $$               %%            %%       ##            ##              ##                  " << endl;
    cout << "                      ##                      ##      @@@@@@@@@@    ##                     ##     @@@@@@@@@@                       @@@@@@@@@@@    @@               @@       @@@@@@@     @@@@@@@@@@@     @@@@@@@@@@@@@@                                  " << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    SetConsoleColor(7);
}

// User Menu:

int UserMenu()
{
    int choice;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << endl;
    cout << "                     1. Go to My Excel" << endl;
    cout << "                     2. Instructions " << endl;
    cout << "                     3. Exit" << endl
         << endl;

    cout << "Enter Your Choice...." << endl;

    while (true)
    {
        cin >> choice;

        if (cin.fail())
        {
            cin.clear();
            cin.ignore(INT_MAX, '\n');
            cout << "Invalid input. Please enter a number: ";
        }
        else if (choice < 1 || choice > 3)
        {
            cout << "Invalid choice. Please enter a number between 1 and 3: ";
        }
        else
        {
            break;
        }
    }
    return choice;
}

// User Guide :

void userGuide()
{
    cout << "                  Excel Operations User Guide:\n";
    cout << "             ----------------------------------------\n";
    cout << "   Arrow Keys:\n";
    cout << "     - Use arrow keys (UP, DOWN, LEFT, RIGHT) for navigation in the Excel sheet.\n";
    cout << "   Insert and Delete Operations:\n";
    cout << "     - 'E': Insert a row above the current cell.\n";
    cout << "     - 'R': Insert a row below the current cell.\n";
    cout << "     - 'W': Insert a column to the left of the current cell.\n";
    cout << "     - 'Q': Insert a column to the right of the current cell.\n";
    cout << "     - 'G': Set a value to the current cell.\n";
    cout << "     - 'D': Delete the current row.\n";
    cout << "     - 'F': Delete the current column.\n";
    cout << "     - 'S': Clear the current column.\n";
    cout << "     - 'A': Clear the current row.\n";
    cout << "  cell Operations:\n";
    cout << "     - 'U': Insert a cell by right shift.\n";
    cout << "     - 'I': Insert a cell by down shift.\n";
    cout << "     - 'J': Delete a cell by left shift.\n";
    cout << "     - 'K': Delete a cell by up shift.\n";
    cout << "  Clipboard Operations:\n";
    cout << "     - 'B': Set the starting cell of the range.\n";
    cout << "     - 'N': Set the ending cell of the range.\n";
    cout << "     - 'C': Copy the range.\n";
    cout << "     - 'X': Cut the range.\n";
    cout << "     - 'V': Paste the range.\n";
    cout << "  Aggregation Operations:\n";
    cout << "     - 'T': Calculate the sum of the cells in the range.\n";
    cout << "     - 'Y': Calculate the average of the cells in the range.\n";
    cout << "     - 'O': Find the maximum value between two cells in the range.\n";
    cout << "     - 'P': Find the minimum value between two cells in the range.\n";
    cout << "     - 'H': Count the number of cells in the range.\n";
    cout << "----------------------------------------\n";
    cout << "Press 'ESC' to exit the program.\n";
}

// cell Class

class cell
{
public:
    string data;
    int row;
    cell *left;
    cell *right;
    cell *up;
    cell *down;

    string color;
    DataType dataType;
    cellAlignment alignment;

    cell(string value, string color, DataType dataType, cellAlignment alignment)
    {
        data = value;
        this->color = color;
        this->dataType = dataType;
        this->alignment = alignment;
        left = nullptr;
        right = nullptr;
        up = nullptr;
        down = nullptr;
    }
    void Select()
    {
        cell *temp = this;

        color = "yellow";
    }

    void Dselect()
    {
        cell *temp = this;
        temp->color = "Green";
    }
};

// Excel Class  :

class Excel
{
private:
    cell *head;
    int totalRows;
    int totalCols;

public:
    Excel()
    {
        start = nullptr;
        current = nullptr;

        totalRows = 5;
        totalCols = 5;
        initializeGrid(totalRows, totalCols, "A");
    }

    int getTotalRows()
    {
        return totalRows;
    }

    int getTotalCols()
    {
        return totalCols;
    }

    // cell Creation function You can create cell by calling this function :

    cell *createcell(string value, string color, DataType dataType, cellAlignment alignment)
    {
        cell *newcell = new cell(value, color, dataType, alignment);

        return newcell;
    }

    // Move Functions that allow you to navigate to any cell in my Excel :

    void MoveRight()
    {
        cell *temp = current;
        if (temp != nullptr && temp->right != nullptr)
        {
            temp->Dselect();
            temp = temp->right;
            temp->Select();
            current = temp; // Update the current cell
        }
    }

    void MoveLeft()
    {
        cell *temp = current;
        if (temp != nullptr && temp->left != nullptr)
        {
            temp->Dselect();
            temp = temp->left;
            temp->Select();
            current = temp;
        }
    }

    void MoveUp()
    {
        cell *temp = current;
        if (temp != nullptr && temp->up != nullptr)
        {
            temp->Dselect();
            temp = temp->up;
            temp->Select();
            current = temp;
        }
    }
    void MoveDown()
    {
        cell *temp = current;
        if (temp != nullptr && temp->down != nullptr)
        {
            temp->Dselect();
            temp = temp->down;
            temp->Select();
            current = temp;
        }
    }

    // Grid of My Excel : I have insert values in each cell of the Excel

    void initializeGrid(int rows, int cols, string value)
    {

        char a = 'A';
        cell *current_cell = createcell(value + to_string(0), "Green", INT_TYPE, LEFT);
        start = current_cell;
        current = current_cell;

        for (int i = 0; i < cols - 1; i++)
        {
            cell *newcell = createcell(char(++a) + to_string(0), "Green", INT_TYPE, LEFT);
            current_cell->right = newcell;
            newcell->left = current_cell;
            current_cell = newcell;
        }

        cell *firstcellInRow = current;
        for (int i = 1; i < rows; i++)
        {

            a = 'A';
            cell *newRowcell = createcell(char(a++) + to_string(i), "Green", INT_TYPE, LEFT);
            current = newRowcell;
            firstcellInRow->down = newRowcell;
            newRowcell->up = firstcellInRow;
            current_cell = newRowcell;

            for (int j = 1; j < cols; j++)
            {
                cell *newcell = createcell(char(a++) + to_string(i), "Green", INT_TYPE, LEFT);
                current_cell->right = newcell;
                newcell->left = current_cell;
                current_cell = newcell;
                firstcellInRow = firstcellInRow->right;
                current_cell->up = firstcellInRow;
                firstcellInRow->down = current_cell;
            }
            firstcellInRow = newRowcell;
        }
    }

    // Display Excel Function : This function is responsible to Display the Grid of My Excel

    void DisplayExcel()
    {
        system("cls");
        cout << endl;
        cout << endl;
        cout << endl;
        cout << endl;
        cout << endl;
        cout << endl;
        cell *currentRow = start;
        while (currentRow != nullptr)
        {
            for (int i = 0; i <= 4; ++i)
            {
                cell *currentCol = currentRow;
                while (currentCol != nullptr)
                {
                    if (i == 0)
                    {
                        if (currentCol->color == "Green")
                        {
                            cout << ANSI_COLOR_GREEN;
                        }
                        else if (currentCol->color == "yellow")
                        {
                            cout << ANSI_COLOR_YELLOW;
                        }

                        cout << "+------------------------";
                    }
                    else
                    {
                        if (currentCol->color == "Green")
                        {
                            cout << ANSI_COLOR_GREEN;
                        }
                        else if (currentCol->color == "yellow")
                        {
                            cout << ANSI_COLOR_YELLOW;
                        }
                        if (i == 2)
                        {
                            cout << "|      " << setw(10) << currentCol->data << "        ";
                        }
                        else
                        {
                            cout << "|                        ";
                        }
                        // cout << "|      " << setw(16) << currentCol->data << "       " << ANSI_COLOR_RESET;
                    }

                    currentCol = currentCol->right;
                }

                if (i == 0)
                {
                    cout << "+" << ANSI_COLOR_RESET << endl;
                }
                else
                {
                    cout << "|" << endl;
                }
            }

            currentRow = currentRow->down;
        }
        cout << "+------------------------+------------------------+------------------------+------------------------+------------------------+" << endl
             << endl;
    }

    // Function to insert a row above the current cell
    void InsertRowAbove()
    {
        cell *newRow = nullptr;
        cell *currentRow = StartcellOFRow(current);
        cell *rowUpCurrent = StartcellOFRow(current)->up;

        for (int i = 0; currentRow != nullptr; i++)
        {
            cell *newRowcell = createcell("  ", " ", INT_TYPE, LEFT);
            newRowcell->left = newRow;
            if (newRow != nullptr)
            {
                newRow->right = newRowcell;
            }
            newRowcell->down = currentRow;
            currentRow->up = newRowcell;
            newRow = newRowcell;
            currentRow = currentRow->right;
        }

        newRow = StartcellOFRow(newRow);
        if (!rowUpCurrent)
        {
            start = newRow;
        }

        for (int i = 0; rowUpCurrent != nullptr; i++)
        {
            newRow->up = rowUpCurrent;
            rowUpCurrent->down = newRow;
            rowUpCurrent = rowUpCurrent->right;
            newRow = newRow->right;
        }
        rows++;
    }
    // Insert a row below the current row
    void InsertRowBelow()
    {
        cell *newRow = nullptr;
        cell *currentRow = StartcellOFRow(current);
        cell *rowBelowCurrent = StartcellOFRow(current)->down;

        for (int i = 0; currentRow != nullptr; i++)
        {
            cell *newRowcell = createcell("  ", " ", INT_TYPE, LEFT);
            newRowcell->left = newRow;
            if (newRow != nullptr)
            {
                newRow->right = newRowcell;
            }
            newRowcell->up = currentRow;
            currentRow->down = newRowcell;
            newRow = newRowcell;
            currentRow = currentRow->right;
        }
        newRow = StartcellOFRow(newRow);
        for (int i = 0; rowBelowCurrent != nullptr; i++)
        {
            newRow->down = rowBelowCurrent;
            rowBelowCurrent->up = newRow;
            newRow = newRow->right;
            rowBelowCurrent = rowBelowCurrent->right;
        }
        rows++;
    }

    // Insert a column to the right of the current column
    void InsertColumnToRight()
    {
        cell *currentColumn = StartcellOFCol(current);

        cell *newCol = nullptr;
        for (int i = 0; currentColumn != nullptr; i++)
        {
            cell *newcell = createcell("  ", " ", INT_TYPE, LEFT);
            newcell->up = newCol;

            if (newCol != nullptr)
                newCol->down = newcell;

            newcell->right = currentColumn->right;
            currentColumn->right = newcell;
            if (newcell->right != nullptr) // edge case
                newcell->right->left = newcell;
            newcell->left = currentColumn;
            newCol = newcell;

            currentColumn = currentColumn->down;
        }

        cols++;
    }

    // Insert a column to the left of the current column
    void InsertColumnToLeft()
    {
        cell *currentColumn = StartcellOFCol(current); // The current column
        cell *newCol = nullptr;                        // New column cell

        for (int i = 0; currentColumn != nullptr; i++)
        {
            cell *newcell = createcell("   ", " ", INT_TYPE, LEFT);
            newcell->up = newCol;

            if (newCol != nullptr)
                newCol->down = newcell;

            newcell->left = currentColumn->left;
            currentColumn->left = newcell;
            if (newcell->left != nullptr)
                newcell->left->right = newcell;
            newcell->right = currentColumn;
            newCol = newcell;

            currentColumn = currentColumn->down;
        }

        cols++;
    }

    // Set the value of the current cell
    void SetcellValue(string value)
    {
        if (value.length() <= 4)
        {
            current->data = value;
        }
        else
        {
            cout << "Enter value again : Value should not be greater than 4 characters." << endl;
        }
    }

    // Delete the current row
    void DeleteRow()
    {
        cell *currentRow = StartcellOFRow(current);

        for (int i = 0; currentRow->up != nullptr && currentRow->down != nullptr; i++)
        {
            currentRow->up->down = currentRow->down;
            currentRow->down->up = currentRow->up;
            currentRow->up = currentRow->up->right;
            currentRow->down = currentRow->down->right;
        }

        if (currentRow->up != nullptr)
        {
            for (int i = 0; currentRow->down != nullptr; i++)
            {
                currentRow->down->up = nullptr;
                currentRow->down = currentRow->down->right;
            }
        }
        if (!currentRow->down)
        {
            for (int i = 0; currentRow->up != nullptr; i++)
            {
                currentRow->up->down = nullptr;
                currentRow->up = currentRow->up->right;
            }
        }

        rows--;
    }

    // Delete the current column
    void DeleteColumn()
    {
        cell *currentCol = StartcellOFCol(current);
        cell *colToLeft = currentCol->left;
        cell *colToRight = currentCol->right;

        for (int i = 0; colToLeft != nullptr && colToRight != nullptr; i++)
        {
            colToLeft->right = colToRight;
            colToRight->left = colToLeft;
            colToLeft = colToLeft->down;
            colToRight = colToRight->down;
        }

        if (!currentCol->left)
        {
            start = colToRight;
            while (colToRight != nullptr)
            {
                colToRight->left = nullptr;
                colToRight = colToRight->down;
            }
        }
        if (!currentCol->right)
        {
            while (colToLeft != nullptr)
            {
                colToLeft->right = nullptr;
                colToLeft = colToRight->down;
            }
        }
        cols--;
    }

    // Function to clear a column

    void ClearColumn()
    {
        cell *currentCol = StartcellOFCol(current);
        while (currentCol != nullptr)
        {
            currentCol->data = "";
            currentCol = currentCol->down;
        }
    }

    // Function to clear a row

    void ClearRow()
    {
        cell *currentRow = StartcellOFRow(current);
        while (currentRow != nullptr)
        {
            currentRow->data = "";
            currentRow = currentRow->right;
        }
    }

    // Function to insert a cell by right shift
    void InsertcellByRightShift()
    {
        cell *newcell = createcell("   ", " ", INT_TYPE, LEFT);
        cell *currentcell = current;
        cell *cellUpCurrent = currentcell->up;
        cell *cellDownCurrent = currentcell->down;
        cell *cellRightCurrent = currentcell->right;
        cell *cellLeftCurrent = currentcell->left;

        if (start == current)
            start = newcell;
        if (!currentcell->right)
        {
            currentcell->right = newcell;
            newcell->left = currentcell;
        }
        else
        {
            current = newcell;
            while (currentcell)
            {
                newcell->right = currentcell;
                if (cellLeftCurrent)
                {
                    currentcell->left = newcell;
                    newcell->left = cellLeftCurrent;
                    cellLeftCurrent->right = newcell;
                    cellLeftCurrent = cellLeftCurrent->right;
                }
                currentcell->left = newcell;
                newcell = currentcell;
                currentcell = currentcell->right;
            }

            newcell = current;
            while (cellUpCurrent)
            {
                newcell->up = cellUpCurrent;
                cellUpCurrent->down = newcell;
                cellUpCurrent = cellUpCurrent->right;
                newcell = newcell->right;
            }
            newcell = current;
            while (cellDownCurrent)
            {
                newcell->down = cellDownCurrent;
                cellDownCurrent->up = newcell;
                cellDownCurrent = cellDownCurrent->right;
                newcell = newcell->right;
            }
        }
        cols++;
    }

    // Function to insert a cell by down shift
    void InsertcellByDownShift()
    {
        cell *newcell = createcell("   ", " ", INT_TYPE, LEFT);
        cell *currentcell = current;
        cell *cellUpCurrent = currentcell->up;
        cell *cellDownCurrent = currentcell->down;
        cell *cellLeftCurrent = currentcell->left;
        cell *cellRightCurrent = currentcell->right;

        if (start == current)
            start = newcell;

        if (!currentcell->down)
        {
            currentcell->down = newcell;
            newcell->up = currentcell;
        }
        else
        {
            current = newcell;

            while (currentcell != nullptr)
            {
                newcell->down = currentcell;

                if (cellUpCurrent != nullptr)
                {
                    currentcell->up = newcell;
                    newcell->up = cellUpCurrent;
                    cellUpCurrent->down = newcell;
                    cellUpCurrent = cellUpCurrent->down;
                }

                currentcell->up = newcell;
                newcell = currentcell;
                currentcell = currentcell->down;
            }

            newcell = current;

            while (cellLeftCurrent != nullptr)
            {
                newcell->left = cellLeftCurrent;
                cellLeftCurrent->right = newcell;
                cellLeftCurrent = cellLeftCurrent->down;
                newcell = newcell->down;
            }

            newcell = current;

            while (cellRightCurrent != nullptr)
            {
                newcell->right = cellRightCurrent;
                cellRightCurrent->left = newcell;
                cellRightCurrent = cellRightCurrent->down;
                newcell = newcell->down;
            }
        }

        rows++;
    }

    // Function to delete the current cell and shift cell leftwards
    void DeletecellByLeftShift()
    {

        if (current->right == nullptr)
        {
            current->data = " ";
        }
        else
        {
            cell *temp = current;
            cell *tempNext = current->right;

            current->data = tempNext->data;
            current->right = tempNext->right;

            delete tempNext;
        }
    }

    // Function to delete the current cell and shift cell upwards
    void DeletecellByUpShift()
    {
        if (current->down == nullptr)
        {

            current->data = " ";
        }
        else
        {
            cell *temp = current;
            cell *tempdown = current->down;

            current->data = tempdown->data;
            tempdown->data = " ";
            current->up = tempdown->up;

            delete tempdown;
        }
    }

    // Function to check whether the range is a row or a column
    bool WhetherRoworColumn(cell *rangeStart, cell *rangeEnd)
    {
        cell *temp = rangeStart;
        while (temp != nullptr)
        {
            if (temp == rangeEnd)
                return true;
            temp = temp->down;
        }
        return false;
    }

    // Function to get the sum of a range
    void GetRangeSum(cell *rangeStart, cell *rangeEnd)
    {
        int sum = 0;

        cell *temp = rangeStart;
        while (temp != nullptr)
        {
            try
            {
                int cellValue = stoi(temp->data);
                sum += cellValue;
            }
            catch (const invalid_argument &e)
            {
            }
            temp = temp->right;
        }

        SetcellValue(to_string(sum));
    }

    // Function to get the average of a range
    void GetRangeAverage(cell *rangeStart, cell *rangeEnd)
    {
        int sum = 0;
        int count = 0;

        cell *temp = rangeStart;
        while (temp != rangeEnd->right)
        {
            try
            {
                int cellValue = stoi(temp->data);
                sum += cellValue;
                count++;
            }
            catch (const invalid_argument &e)
            {
            }
            temp = temp->right;
        }

        double result = static_cast<double>(sum) / count;
        SetcellValue(to_string(result));
    }

    // Function to get the count of a range
    void GetRangeCount(cell *rangeStart, cell *rangeEnd)
    {
        int count = 0;

        cell *temp = rangeStart;
        while (temp != rangeEnd->right)
        {
            try
            {
                count++;
            }
            catch (const invalid_argument &e)
            {
            }
            temp = temp->right;
        }
        SetcellValue(to_string(count));
    }

    // Function to get the minimum value in a range
    void GetRangeMin(cell *rangeStart, cell *rangeEnd)
    {
        int minVal = INT_MAX;

        cell *temp = rangeStart;

        while (temp != nullptr && temp != rangeEnd->right)
        {
            try
            {
                int cellValue = stoi(temp->data);
                if (cellValue < minVal)
                {
                    minVal = cellValue;
                }
            }
            catch (const invalid_argument &e)
            {
            }
            temp = temp->right;
        }
        SetcellValue(to_string(minVal));
    }

    // Function to get the maximum value in a range
    void GetRangeMax(cell *rangeStart, cell *rangeEnd)
    {
        int maxVal = INT_MIN;

        cell *temp = rangeStart;

        while (temp != nullptr && temp != rangeEnd->right)
        {
            try
            {
                int cellValue = stoi(temp->data);
                if (cellValue > maxVal)
                {
                    maxVal = cellValue;
                }
            }
            catch (const invalid_argument &e)
            {
            }
            temp = temp->right;
        }
        SetcellValue(to_string(maxVal));
    }

    // Function to copy a range of cells
    vector<string> CopyRange(cell *rangeStart, cell *rangeEnd, vector<string> CopyorCut)
    {
        bool column = WhetherRoworColumn(rangeStart, rangeEnd);

        if (!column)
        {
            cell *currentcell = rangeStart;
            while (currentcell != nullptr && currentcell != rangeEnd->right)
            {
                CopyorCut.push_back(currentcell->data);
                currentcell = currentcell->right;
            }
        }
        else
        {
            cell *currentcell = rangeStart;
            while (currentcell != nullptr && currentcell != rangeEnd->down)
            {
                CopyorCut.push_back(currentcell->data);
                currentcell = currentcell->down;
            }
        }
        return CopyorCut;
    }

    // Function to cut a range of cells
    vector<string> CutRange(cell *rangeStart, cell *rangeEnd, vector<string> CopyorCut)
    {
        bool column = WhetherRoworColumn(rangeStart, rangeEnd);

        if (!column)
        {
            cell *currentcell = rangeStart;
            while (currentcell != nullptr && currentcell != rangeEnd->right)
            {
                CopyorCut.push_back(currentcell->data);
                currentcell->data = " ";
                currentcell = currentcell->right;
            }
        }
        else
        {
            cell *currentcell = rangeStart;
            while (currentcell != nullptr && currentcell != rangeEnd->down)
            {
                CopyorCut.push_back(currentcell->data);
                currentcell->data = " ";
                currentcell = currentcell->down;
            }
        }

        return CopyorCut;
    }

    // Function to paste copied/cut data to a range

    void Paste(const vector<string> &CopyorCut, cell *rangeStart, cell *rangeEnd)
    {

        bool column = WhetherRoworColumn(rangeStart, rangeEnd);

        if (!column)
        {
            cell *startcell = current;

            for (const string &data : CopyorCut)
            {
                startcell->data = data;

                if (startcell->right != nullptr)
                {
                    startcell = startcell->right;
                }
                else
                {

                    cell *newcell = createcell(" ", " ", INT_TYPE, LEFT);
                    newcell->left = startcell;
                    startcell->right = newcell;
                    startcell = newcell;
                }
            }
        }
        else
        {
            cell *startcell = current;

            for (const string &data : CopyorCut)
            {
                startcell->data = data;

                if (startcell->down != nullptr)
                {
                    startcell = startcell->down;
                }
                else
                {

                    cell *newcell = createcell(" ", " ", INT_TYPE, LEFT);
                    newcell->up = startcell;
                    startcell->down = newcell;
                    startcell = newcell;
                }
            }
        }
    }

    // Function to save data to the File Name ExcelData.txt

    void SaveDataToTheFile(const string &ExcelData)
    {
        ofstream outfile(ExcelData);
        if (outfile.is_open())
        {
            cell *currentRow = start;
            while (currentRow != nullptr)
            {
                cell *currentCol = currentRow;
                while (currentCol != nullptr)
                {
                    outfile << currentCol->data << ",";
                    currentCol = currentCol->right;
                }
                outfile << endl;
                currentRow = currentRow->down;
            }
            outfile.close();
        }
    }

    // Load data from the FIle :

    void LoadDataFromFile(const string &filename)
    {
        ifstream file(filename);

        if (file.is_open())
        {
            string line;
            cell *currentRow = start;
            int row = 0;

            while (getline(file, line))
            {
                istringstream iss(line);
                string value;

                int col = 0;
                while (getline(iss, value, ','))
                {

                    cell *currentcell = getcell(row, col);
                    if (currentcell != nullptr)
                    {
                        currentcell->data = value;
                    }

                    col++;
                }

                row++;
                currentRow = currentRow->down;
            }

            file.close();
        }
    }

    // Get any cell of the Excel through this function :

    cell *getcell(int rowIndex, int columnIndex)
    {
        cell *startcell = start;

        for (int i = 0; i < rowIndex; i++)
        {
            startcell = startcell->down;
        }
        for (int j = 0; j < columnIndex; j++)
        {
            startcell = startcell->right;
        }

        return startcell;
    }

    // Through this function you can set the value of any cell in My Excel :

    void SetCurrent(cell *current)
    {
        this->current = current;
    }
    cell *getcurrentcell()
    {
        return current;
    }

private:
    cell *start;
    cell *current;
    int rows;
    int cols;

    cell *StartcellOFRow(cell *current)
    {
        cell *currentcell = current;
        while (currentcell->left)
        {
            currentcell = currentcell->left;
        }
        return currentcell;
    }

    cell *StartcellOFCol(cell *current)
    {
        cell *currentcell = current;
        while (currentcell->up)
        {
            currentcell = currentcell->up;
        }
        return currentcell;
    }
};

// Excel OPerations :

void ExcelOperations()
{
    Excel obj1;
    vector<string> CopyorCut;
    // obj1.LoadDataFromFile("ExcelData.txt");ma
    obj1.DisplayExcel();
    cell *rangeStart = nullptr;
    cell *rangeEnd = nullptr;
    while (true)
    {

        // Arrow keys for navigation
        if (GetAsyncKeyState(VK_RIGHT))
        {
            obj1.MoveRight();
            obj1.DisplayExcel();
        }
        else if (GetAsyncKeyState(VK_LEFT))
        {
            obj1.MoveLeft();
            obj1.DisplayExcel();
        }
        else if (GetAsyncKeyState(VK_UP))
        {
            obj1.MoveUp();
            obj1.DisplayExcel();
        }
        else if (GetAsyncKeyState(VK_DOWN))
        {
            obj1.MoveDown();
            obj1.DisplayExcel();
        }

        // 'E' key for inserting a row above
        else if (GetAsyncKeyState(0x45))
        {
            obj1.InsertRowAbove();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'R' key for inserting a row below
        else if (GetAsyncKeyState(0x52))
        {
            obj1.InsertRowBelow();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'W' key for inserting a column to the left
        else if (GetAsyncKeyState(0x57))
        {
            obj1.InsertColumnToLeft();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'Q' key for inserting a column to the right
        else if (GetAsyncKeyState(0x51))
        {
            obj1.InsertColumnToRight();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'G' key for setting a value to the current cell
        else if (GetAsyncKeyState(0x47))
        {
            string val;
            cout << "Enter Value of the cell" << endl;
            cin >> val;
            obj1.SetcellValue(val);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'D' key for deleting a row
        else if (GetAsyncKeyState(0x44))
        {
            obj1.DeleteRow();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'F' key for deleting a column
        else if (GetAsyncKeyState(0x46))
        {
            obj1.DeleteColumn();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'S' key for clearing a column
        else if (GetAsyncKeyState(0x53))
        {
            obj1.ClearColumn();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'A' key for clearing a row
        else if (GetAsyncKeyState(0x41))
        {
            obj1.ClearRow();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // 'U' key for inserting a cell by right shift
        else if (GetAsyncKeyState(0x55))
        {
            obj1.InsertcellByRightShift();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // 'I' key for inserting a cell by down shift
        else if (GetAsyncKeyState(0x49))
        {
            obj1.InsertcellByDownShift();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // 'J' key for deleting a cell by left shift
        else if (GetAsyncKeyState(0x4A))
        {
            obj1.DeletecellByLeftShift();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // 'K' key for deleting a cell by up shift
        else if (GetAsyncKeyState(0x4B))
        {
            obj1.DeletecellByUpShift();
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // To Get the value of the rangeStart ('B' key)
        if (GetAsyncKeyState(0x42) & 0x8000)
        {

            rangeStart = obj1.getcurrentcell();
        }

        // To get the value of the rangeEnd ('N' key)
        if (GetAsyncKeyState(0x4E) & 0x8000)
        {

            rangeEnd = obj1.getcurrentcell();
        }

        // 'C' key for copying a range
        if (GetAsyncKeyState(0x43) & 0x8000)
        {
            CopyorCut = obj1.CopyRange(rangeStart, rangeEnd, CopyorCut);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // 'X' key for cutting a range
        if (GetAsyncKeyState(0x58) & 0x8000)
        {
            CopyorCut = obj1.CutRange(rangeStart, rangeEnd, CopyorCut);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }

        // 'V' key for pasting a range
        else if (GetAsyncKeyState(0x56) & 0x8000)
        {
            obj1.Paste(CopyorCut, rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // T Key For the SUm of the cells
        else if (GetAsyncKeyState(0x54) & 0x8000)
        {
            obj1.GetRangeSum(rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // Y key for the Average of the cells
        else if (GetAsyncKeyState(0x59) & 0x8000)
        {
            obj1.GetRangeAverage(rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // O key for the Maximum value between 2 cells
        else if (GetAsyncKeyState(0x4F) & 0x8000)
        {
            obj1.GetRangeMax(rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // P key for the Minimum value between 2 cells
        else if (GetAsyncKeyState(0x50) & 0x8000)
        {
            obj1.GetRangeMin(rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        // H key for the Count of the  cells
        else if (GetAsyncKeyState(0x48) & 0x8000)
        {
            cout << "H key is presseed";
            obj1.GetRangeCount(rangeStart, rangeEnd);
            obj1.DisplayExcel();
            obj1.SaveDataToTheFile("ExcelData.txt");
        }
        else if (GetAsyncKeyState(0x1B) & 0x8000)
        {
            return;
        }

        Sleep(50);
    }
}
int main()
{
    Header();

    int choice;
    do
    {
        choice = UserMenu();
        if (choice == 1)
        {
            ExcelOperations();
            Header();
        }
        else if (choice == 2)
        {
            system("cls");

            userGuide();
            while (_getch() != 27)
            {
            }

            system("cls");
            Header();
        }
    } while (choice != 3);

    return 0;
}
