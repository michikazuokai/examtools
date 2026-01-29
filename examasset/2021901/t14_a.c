int flg = -1;

for (int i = 0; i < 5; i++){
    if (a[i] == v){
        flg = 0;
        break;
    }
}

if (flg == 0) {
    printf("みつかった");
} else {
    printf("見つからなかった");
}