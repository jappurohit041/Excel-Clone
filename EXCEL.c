/*
	It is very basic version of Microsoft Excel which has following functionalities:
	a) New file option (creates a new file)
	b) Saving a file (to .csv format)
	c) Open a file (which is in .csv format)
	d) Graph Plotting (Graph plotting with labels)
	e) Average (find the average of the data set or particular column/row)
	f) Sum (find the total sum of the data set or particular column/row)
	g) Sort (sort a particular data in asscending order)

*/
#include<stdio.h>
#include<string.h>
#include<windows.h>
#include<limits.h>
#include<ctype.h>

#define NO_COL 7
#define NO_ROW 24
#define MAX_COL 64
#define MAX_ROW 256
#define NO_HDR 3

#define SCN_BAR 2
#define OFF_SET 4

#define cprintf printf
#define CELL_MEM_SIZE 60

char *file[MAX_ROW+1][MAX_COL+1] = {0};

int getch();
int getche();

struct
{
	int x, y;
} fream = {0, 0}, cursor = {0, 0};

void setcolor(int color) { SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color); }
void setcolor_bright() { SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), 0x70); }
void setcolor_dark() { SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), 0x07); }
void clrscr() { system("cls"); }

// print value in header from cell
char* get_header(char str[], int row, int col)
{
	if(file[row][col] != 0)
		strncpy(str, file[row][col], CELL_MEM_SIZE);
	else
		strncpy(str, "", CELL_MEM_SIZE);

	return str;
}
// geting value from the cell
char* get_val(char str[], int row, int col)
{
	if(file[row][col] != 0)
	{
		strncpy(str, file[row][col], 10);
		str[10] = 0;
	}
	else
	{
		strncpy(str, "", 10);
	}

	return str;
}
// get column name
char* get_col(char str[], int val)
{
	if(val < 26)
	{
		str[0] = 'A' + val;
		str[1] = 0;
	}
	else
	{
		str[0] = 'A' + val/26 - 1;
		str[1] = 'A' + val%26;
		str[2] = 0;
	}

	return str;
}
// save a file
void save()
{
	char fname[30] = "";
	char str[60] = "";
	int i, j;
	FILE *fp;

	clrscr();

	printf("Enter name of file: ");
	gets(fname);

	strcat(fname, ".csv");

	fp = fopen(fname, "w");

	for(i = 0; i < MAX_ROW; i++)
	{
		for(j = 0; j < MAX_COL; j++)
		{
			fputs(get_header(str, i, j), fp);
			if(j != MAX_COL-1)
			{
				fputs(",", fp);
			}
		}
		fputs("\n", fp);
	}

	fclose(fp);
}
// free file that is used for creating new file
void free_file(){
        int i, j;
	for(i = 0; i < MAX_ROW; i++)
	{
		for(j = 0; j < MAX_COL; j++)
		{
		        if(file[i][j] != 0) free(file[i][j]);
		        file[i][j] = 0;
		}
        }
}
void put_file(int i, int j, char* str){
        if(strcmp(str, "") != 0
           && 0 <= i && i < MAX_ROW
           && 0 <= j && j < MAX_COL) {
                file[i][j] = (char*)malloc(sizeof(char) * CELL_MEM_SIZE);
                strcpy(file[i][j], str);
        }
}
// open csv file
void open()
{
	char fname[30] = "";
	char str[CELL_MEM_SIZE ] = "";
	int i, j;
	int k;
	int ch;
	FILE *fp;

	clrscr();

	printf("Enter name of file: ");
	gets(fname);

	strcat(fname, ".csv");

	fp = fopen(fname, "r");

	if(fp == NULL){
                printf("Error opening file");
                getch();
                return;
	}

	i = j = k = 0;
	free_file();
	while(1){
                ch = getc(fp);

                if(ch == EOF || ch == ',' || ch == '\n')
                {
                        str[k] = 0;
                        put_file(i, j, str);
                        strcpy(str, "");
                }

                if(ch == EOF){
                        break;
                }
                else if(ch == ','){
                        k = 0;
                        j++;
                }
                else if(ch == '\n'){
                        j = k = 0;
                        i++;
                }
                else{
                        str[k] = ch;
                        k++;
                }
	}

	fclose(fp);
}
// display the structure of excel
void show()
{
	int i, j;
	char str[70] = "";

	int si = fream.y, ei = fream.y + NO_ROW;
	int sj = fream.x, ej = fream.x + NO_COL;

	setcolor_dark();
	system("cls");

	setcolor(0xE4);
	printf("%-76s ", "F1-NEW  F2-SAVE  F3-OPEN F4-GRAPH");
        printf("\n");

        // setcolor(0x9E);
        setcolor(0x70);
	cprintf("%-76s ", get_header(str, cursor.y, cursor.x) );
	printf("\n\n");

	setcolor_bright();
	cprintf("%6s", "");
	for(j = sj; j < ej; j++)
	{
		setcolor_bright();
		printf("%6s", get_col(str, j));
		printf("%4s", "");
	}
        setcolor_bright();
        printf(" ");

	printf("\n");
	for(i = si; i < ei; i++)
	{
		setcolor_bright();
		cprintf("%3d | ", i+1);

		for(j = sj; j < ej; j++)
		{
			if(j == cursor.x && i == cursor.y)
			{
                setcolor_bright();
			}
			else
			{
                setcolor_dark();
			}
			get_val(str, i, j);
			cprintf("%10s", str);
		}
        setcolor_bright();
        printf(" ");
		printf("\n");
	}

	setcolor(0xE4);
	printf("%76s ", "F5-AVG  F6-SUM  F7-SORT");
	printf("\n");

	setcolor_dark();
}
// string to int for column
int col_to_idx(char str[]) {
        int col = 0;
        int len = strlen(str);
        strupr(str);

        if(len > 2 || len == 0) {
                return -1;
        }
        else if(len == 2){
                if(str[0] > 'Z' || str[0] < 'A' || str[1] > 'Z' || str[1] < 'A'){
                        return -1;
                }
                col = (str[0]-'A' + 1)*26 + (str[1]-'A');
        }
        else if(len == 1){
                if(str[0] > 'Z' || str[0] < 'A'){
                        return -1;
                }
                col = (str[0]-'A');
        }

        return col;
}
//string to int for row
int row_to_idx(char str[]) { return atoi(str)-1; }

// ploting graph
void graph_it(){
        char str[CELL_MEM_SIZE] = "";

        int si = 0, ei = 4;
        int sj = 0, ej = 2;

        int max =0;
        int min =0;

        int i, j;

        int value;
        int y;

        clrscr();

        printf("Enter start row: ");
        scanf("%s", str);
        si = row_to_idx(str);
        printf("Enter start col: ");
        scanf("%s", str);
        sj = col_to_idx(str);

        printf("Enter end row: ");
        scanf("%s", str);
        ei = row_to_idx(str);
        printf("Enter end col: ");
        scanf("%s", str);
        ej = col_to_idx(str);

        for(i = si+1; i <= ei; i++){
                for(j = sj+1; j <= ej; j++){
                        value = atoi(file[i][j]);
                        if(max < value) { max = value; }
                        if(min > value) { min = value; }
                }
        }

        int diff = max-min;
        int scale = 0;

        if(diff > 23){
                scale = (max-min)/23;
        }
        else{
                scale = 1;
        }
        for(y = max; y >= min; y -= scale) {
                printf("%5d ", y);
                for(i = si+1; i <= ei; i++){
                        for(j = sj+1; j <= ej; j++){
                                value = atoi(file[i][j]);
                                if(value >= y){
                                        printf("%c%c", 219, 219);
                                }
                                else{
                                        printf("%2s", "");
                                }
                                printf("%2s", "");
                        }
                        printf("%4s", "");
                }
                printf("\n");
        }

        printf("%6s", "");
        for(i = si+1; i <= ei; i++){
                int diff_j = ej-sj;
                int width = diff_j * 4 + 4;
                strncpy(str, file[i][sj], width);
                str[min(width, 59)] = 0;

                printf("%-*s", width, str);
        }

        getch();
}
// average function
void insert_avg(){
        char str[5] = "";

        int si = 0, ei = 4;
        int sj = 0, ej = 2;

        int i, j;

        int value;

        clrscr();

        printf("Enter start row: ");
        scanf("%s", str);
        si = row_to_idx(str);
        printf("Enter start col: ");
        scanf("%s", str);
        sj = col_to_idx(str);

        printf("Enter end row: ");
        scanf("%s", str);
        ei = row_to_idx(str);
        printf("Enter end col: ");
        scanf("%s", str);
        ej = col_to_idx(str);

        int sum = 0;

        for(i = si; i <= ei; i++){
                for(j = sj; j <= ej; j++){
                        value = atoi(file[i][j]);
                        sum += value;
                }
        }

        int n = (ei-si + 1)*(ej-sj + 1);

        if(n > 0){
                int avg = sum/n;
                put_file(cursor.y, cursor.x, itoa(avg, str, 10));
        }
}
// sum function
void insert_sum(){
        char str[5] = "";

        int si = 0, ei = 4;
        int sj = 0, ej = 2;

        int i, j;

        int value;

        clrscr();

        printf("Enter start row: ");
        scanf("%s", str);
        si = row_to_idx(str);
        printf("Enter start col: ");
        scanf("%s", str);
        sj = col_to_idx(str);

        printf("Enter end row: ");
        scanf("%s", str);
        ei = row_to_idx(str);
        printf("Enter end col: ");
        scanf("%s", str);
        ej = col_to_idx(str);

        int sum = 0;

        for(i = si; i <= ei; i++){
                for(j = sj; j <= ej; j++){
                        value = atoi(file[i][j]);
                        sum += value;
                }
        }

        int n = (ei-si + 1)*(ej-sj + 1);

        if(n > 0){
                put_file(cursor.y, cursor.x, itoa(sum, str, 10));
        }
}

// comparing to excel cell
int excel_strcmp(char *p1, char *p2){
        if(p1 == 0){
                return 1;
        }
        else if(p2 == 0){
                return -1;
        }
        else {
                return strcmp(p1, p2);
        }
}
//sorting column function
void sort_col(){
        char str[30];
        char *temp[MAX_ROW];

        printf("Enter col: ");
        scanf("%s", str);

        int col = col_to_idx(str);

        for(int i = 0; i < MAX_ROW; i++){
                for(int j = 0; j < MAX_ROW; j++){
                        if(excel_strcmp(file[i][col], file[j][col]) < 0){
                                memcpy(temp, file[i], sizeof(file[i]));
                                memcpy(file[i], file[j], sizeof(file[i]));
                                memcpy(file[j], temp, sizeof(file[i]));
                        }
                }
        }
}
// printing cursor
void print_cur(char *s)
{
	int x, y;

	x = 10*(cursor.x-fream.x) + 6;
	y = (cursor.y-fream.y) + OFF_SET;
	gotoxy(x, y);
	setcolor_bright();
	cprintf("%10s", s);
}
// taking input
void input()
{
	int diff;
	int ch, i;
	char *temp;
	ch = getch();

	if(isalnum(ch))
	{
		if(file[cursor.y][cursor.x] == 0)
		{
			file[cursor.y][cursor.x] = (char*)calloc(sizeof(char), CELL_MEM_SIZE);
		}
		temp = file[cursor.y][cursor.x];

		memset(temp, 0, CELL_MEM_SIZE);

		temp[0] = ch;
		i = 1;

		gotoxy(0, SCN_BAR);
		printf("%c", ch);
		print_cur(temp);

		while( setcolor_dark(), gotoxy(i, SCN_BAR), (ch = getche()) != '\r' && (isalnum(ch) || ch == '\b') )
		{
		        if(ch == '\b'){
                                if(i > 0){ printf(" ");  i--; temp[i] = 0; }
                                print_cur(temp);
                                continue;
		        }
			temp[i] = ch;
			print_cur(temp);
			i++;
		}
	}

	switch(ch)
	{
		case 26: clrscr(); exit(0);

		case 0:
                        switch(getch()){
                                case 59: free_file(); break;
                                case 60: save(); break;
                                case 61: open(); break;
                                case 62: graph_it(); break;
                                case 63: insert_avg(); break;
                                case 64: insert_sum(); break;
                                case 65: sort_col(); break;
                        }
                        break;
		case 224:
			/*

				f1: 0 59
				f2: 0 60
				f3: 0 61
				f4: 0 62
				f5: 0 63
				f6: 0 64
				f7: 0 65
				f8: 0 66
				f9: 0 67
				f10: 0 68

			*/

			switch(getch())
			{
				case 72:
					if(cursor.y > 0) cursor.y--;
					diff = cursor.y-fream.y;
					if(diff < 0)
					{
						fream.y--;
					}
					break;
				case 80:
					if(cursor.y < MAX_ROW-1) cursor.y++;
					diff = cursor.y-fream.y;
					if(diff >= NO_ROW)
					{
						fream.y++;
					}
					break;
				case 77:
					if(cursor.x < MAX_COL-1) cursor.x++;
					diff = cursor.x-fream.x;
					if(diff >= NO_COL)
					{
						fream.x++;
					}
					break;
				case 75:
					if(cursor.x > 0) cursor.x--;
					diff = cursor.x-fream.x;
					if(diff < 0)
					{
						fream.x--;
					}
					break;
			}
	}
}
//main function
int main()
{
	while(1)
	{
		show();
		input();
	}
}
