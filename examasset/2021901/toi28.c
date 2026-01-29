#include <stdio.h>

int fsum({%box(①,15)%}){
  int result = p1 + p2;
  return {%box(②,10)%};
}

int main(void)
{
  int data1 = 100;
  int data2 = 200;

  int ans = fsum({%box(③,15)%});
  printf("関数の結果=%d\n",ans);

  return 0;
}