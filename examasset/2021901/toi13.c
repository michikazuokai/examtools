#include <stdio.h>
int main(void){
	int num;
	printf("果物が表示します。数字を入力して下さい");
	scanf("%d",&num);

	switch (num){
		case 1:
			printf("Apple ");
			break;
		case 2:
			printf("Banana ");
		case 3:
			printf("Orange ");
			break;
		default:
			printf("Cherry ");
			break;
	}
	return 0;
}
