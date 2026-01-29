int i;

for (i = 0; i < 5; i++){
    if (a[i] == v){
	    break;
	}
}

if (i < 5){
    printf("みつかった");
} else {
    printf("見つからなかった");
}