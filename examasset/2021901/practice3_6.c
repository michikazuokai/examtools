#include <stdio.h>

int main(void){
    int counts[] = {1, 3, 5, 6, 5, 2};
    for (int i = 0; i < 6 ; i++){
        printf("%d :",i);
        for (int j = j =0; j < counts[i]; j++){
            printf("*");
        }
        printf("\n");
    }
    return 0;
}