#include <stdio.h>
int main(void){
	int i = 0;
	while (i < 5){
		if (i % 2 ==0){
			i ++;
			continue;
		}
		printf("%d ",i);
		i++;
	}
	return 0;
}