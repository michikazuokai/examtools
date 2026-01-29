void showResult(int person, int total, char item[]);

int main(void) {
    int orders[3][4] = {
        {2, 1, 1, 0},
        {1, 2, 2, 1},
        {3, 0, 0, 2}
    };
    int prices[4] = {500, 250, 300, 400};
    char items[4][10] = {"パン", "飲み物", "弁当", "お菓子"};

    for (int i = 0; i < 3; i++) {
        int total = calcTotal(orders[i], prices, 4);
        int maxIndex = findMaxQtyIndex(orders[i], 4);
        {%box(③,15)%} ;
    }

    return 0;
}