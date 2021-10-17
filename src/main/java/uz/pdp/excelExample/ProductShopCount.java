package uz.pdp.excelExample;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ProductShopCount {
    private String shopName;
    private Integer amount;
}
