package uz.pdp.excelExample;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ProductCount {
    private String name;
    private String imgUrl;
    private List<ProductShopCount> productShopCountList;
}
