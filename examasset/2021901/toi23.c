#include <stdio.h>

int main(void){
    int array[] = {15, 5, 0, 22, 8, 13, 14, 71, 49, 7};
    int data;
    int n = -1;    /*  見つかった時のインデックスを保管 */

    // 探索するデータを入力
    printf("数字を入力して下さい");
    scanf("%d", &data);

    // 配列arrayのデータを探す
    for (int i = 0; i <  {%box(①,15)%}); i++ ){
        if (array[i] == data){
{%box(②,10,2,6)%}
        }
    }

    // 結果を表示する
    if ({%box(③,10)%}) {
        printf("%dは%d番目に見つかりました¥n", data, n+1);
    } else {
        printf("%dは見つかりませんでした¥n", data);
    }
    return 0;
};