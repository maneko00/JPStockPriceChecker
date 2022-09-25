export class JpStockGetter {

    private prefix_url_: string = 'https://www.google.com/finance/quote/';
    private suffix_url_: string = ':TYO?hl=ja';
    private html_: string = '';

    private name_: string = '';
    private price_: number = 0;
    private dividends_per_share_: string|number = '';

    public constructor(code: number) {
        // undefinedの代わりに2124を代入
        const code_number = code ?? 2124;
        const url = this.prefix_url_ + code_number.toString() + this.suffix_url_;        
        this.html_ = UrlFetchApp.fetch(url).getContentText("UTF-8");
        
        this.name_ = this.GetCompanyName();
        this.price_ = this.GetStockPrice();
        this.dividends_per_share_ = this.GetDividendsPerShare();
    }

    get name() {
        return this.name_;
    }

    get price() {
        return this.price_;
    }

    get dividends_per_share() {
        return this.dividends_per_share_;
    }

    /**
     * 株情報  
     *
     * @return 株情報
     * @date 2022/9/24
     */
    private GetStockString(from_str: string, to_str: string):string {
        return Parser.data(this.html_)
            .from(from_str)
            .to(to_str)
            .build();
    }

    /**
     * 企業名取得  
     *
     * @return 企業名
     * @date 2022/9/24
     */
    private GetCompanyName():string {
        return this.GetStockString('<div role=\"heading\" aria-level=\"1\" class=\"zzDege\">', '</div>');
    }

    /**
     * 株価取得 
     * 
     * @return 株価
     * @date 2022/9/24
     */
    private GetStockPrice(): number {
        let stock_price_str: string = this.GetStockString('<div class=\"YMlKec fxKbKc\">', '</div>');
        stock_price_str = stock_price_str.replace('￥', '');
        stock_price_str = stock_price_str.replace(',', '');
        return Number(stock_price_str);
    }
    
    /**
     * 配当利回り取得 
     * 
     * @return 1株あたり配当金
     * @date 2022/9/24
     */
    private GetDividendYield():number|string {
        let reference_index_string_array = Parser.data(this.html_)
          .from('<div class=\"P6K39c\">')
          .to('</div>')
          .iterate();

        if (reference_index_string_array.length <= 6) {
            return 0;
        }


      return reference_index_string_array[5].replace('%', '');
    }
    
    /**
     * 1株あたり配当金取得 
     * 
     * @return 1株あたり配当金
     * @date 2022/9/24
     */
    private GetDividendsPerShare():number|string {
        let dividend_yield: number|string = this.GetDividendYield();
        if (dividend_yield == '-') {
            return dividend_yield;
        }
        dividend_yield = Number(dividend_yield);
        
        let stock_price: number = this.GetStockPrice();
        
        return  Math.round((dividend_yield / 100) * stock_price);
    }
}