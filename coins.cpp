#include <libxl.h>
#include <iostream>
#include <locale>
#include <vector>
#include <limits>
using namespace libxl;

bool xlsLoad(const char* fn, std::vector<std::vector<int>> &vec)
{
    Book* xl = xlCreateBook();
    if(!xl->load(fn, NULL))
    {
        return false;
    }

    Sheet *sh = xl->getSheet(0);
    if(!sh) return 1;

    int col,row;
    col = sh->lastFilledCol() - sh->firstFilledCol();
    row = sh->lastFilledRow() - sh->firstFilledRow()+1;

    vec.clear();
    vec.resize(col);
    vec.reserve(col);
    for(auto it = vec.begin(); it != vec.end(); it++)
    {
        it->resize(row);
        it->reserve(row);
    }
    
    for(int i = sh->firstFilledCol(); i < sh->lastFilledCol(); i++)
    {
        for(int j = sh->firstFilledRow()-1; j < sh->lastFilledRow(); j++)
        {
            vec[i][j] = sh->readNum(i,j);
        }
    }

    xl->release();

    return true;
}


class Robot
{
    int _x,_y;
public:
    Robot(int x, int y) : _x(x),_y(y) {}
    Robot() : Robot(0,0) {}
    Robot(const Robot &r) : Robot(r._x,r._y) { }

    void right() { _x++; }
    void up() { _y--; }

    void reset(int x, int y) { _x = x; _y = y; }
    void reset() { reset(0,0); }
    int x() { return _x; }
    int y() { return _y; };
};

struct Answer
{
    int min,max;
    Answer(int mi, int ma) : min(mi), max(ma) { }
    Answer() : Answer(0,0) { }
    Answer(const Answer &a) : Answer(a.min,a.max) { }
};

std::ostream &operator<<(std::ostream &os, const Answer &a)
{
    return os << "Min: " << a.min << "\tMax: " << a.max;
}

void print_vec(const std::vector<std::vector<int>> &field)
{
    for (auto &&i : field)
    {
        for (auto &&j : i)
        {
            std::cout << j << '\t';
        }
        std::cout << '\n';
    }
}

Answer algorythm(const std::vector<std::vector<int>> &field)
{
    Robot r(0, field[0].size()-1);
    Answer ans(field[r.y()][r.x()],field[r.y()][r.x()]);

    //For max
    std::cout << "Field:\n";
    print_vec(field);

    std::vector<std::vector<int>> vec = field;
    vec[r.y()][r.x()] = 0;

    while(!((r.y() <= 0) && (r.x() >= (field[r.y()].size()-1))))
    {
        if(r.y() <= 0)
        {
            r.right();
        }
        else if(r.x() >= field[r.y()].size()-1)
        {
            r.up();
        }
        else
        {
            if(field[r.y()][r.x()+1] < field[r.y()-1][r.x()])
                r.up();
            else
                r.right();
        }
        vec[r.y()][r.x()] = 0;
        ans.max += field[r.y()][r.x()];
    }


    std::cout << "\n\nMax:\n";
    print_vec(vec);

    //For min
    vec = field;

    r.reset(0, field.size()-1);
    vec[r.y()][r.x()] = 0;
    while(!((r.y() <= 0) && (r.x() >= (field[r.y()].size()-1))))
    {
        if(r.y() <= 0)
        {
            r.right();
        }
        else if(r.x() >= field[r.y()].size())
        {
            r.up();
        }
        else
        {
            if(field[r.y()][r.x()+1] > field[r.y()-1][r.x()])
                r.up();
            else
                r.right();
        }
        vec[r.y()][r.x()] = 0;
        ans.min += field[r.y()][r.x()];
    }
    std::cout << "\n\nMin:\n";
    print_vec(vec);

    return ans;
}

int main(int argc, const char **argv)
{
    setlocale(LC_ALL,"");
    if(argc < 2) return 1;

    std::vector<std::vector<int>> field;
    if(!xlsLoad(argv[1], field)) return 1;

    std::cout << algorythm(field) << '\n';
    
    std::cin.clear();
    std::cin.ignore(32767,'\n');
    std::cin.get();

    return 0;
}