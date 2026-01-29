#include <stdio.h>
int main(void){
    int x;
    
    scanf("%d",&x);
    
    if (x < 5){
        printf("Bob");
    } else if (x < 8) {
        printf("Joe");
    } else {
        printf("Tom");
    }
    return 0;
}