package com.generator.excel.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;

@Getter
@Setter
@AllArgsConstructor
public class LogModel {
    private long logId;
    private String name;
    private BigDecimal premLimit;
    private BigDecimal rate;
}
