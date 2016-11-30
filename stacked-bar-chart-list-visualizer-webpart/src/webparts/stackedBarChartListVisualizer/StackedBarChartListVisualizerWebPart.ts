import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './StackedBarChartListVisualizer.module.scss';
import * as strings from 'stackedBarChartListVisualizerStrings';
import { IStackedBarChartListVisualizerWebPartProps } from './IStackedBarChartListVisualizerWebPartProps';
import * as d31 from 'd3';

interface ISPListItem {
  Title: string;
  Month: string;
  Count: string;
}

interface ISPListItems {
  value: ISPListItem[]
}

class ChartDataItem {
  total: number;
  Category: string;
  [s: string]: string | number;
}

class ChartData extends Array<ChartDataItem> {
  columns: string[]
}

export default class StackedBarChartListVisualizerWebPart extends BaseClientSideWebPart<IStackedBarChartListVisualizerWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.stackedBarChartListVisualizer}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Chart based on data from SharePoint List:</span>
              <svg width="400" height="500" id="chart"></svg>
            </div>
          </div>
        </div>
      </div>`;

    this._fetchData();
  }

  private _fetchData(): void {
    this._getData()
      .then((response) => {
        this._renderData(response.value);
      });
  }

  private _getData(): Promise<ISPListItems> {
    return this.context.httpClient
      .get(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists/GetByTitle('Destination%20List')/items?$select=Month,Title,Count")
      .then((response: Response) => {
        return response.json();
      });
  }

  private _renderData(items: ISPListItem[]): void {
    var transformed = this._transform(items);
    this._render(transformed);
  }

  private _transform(data: ISPListItem[]) {
    let result: ChartData = new ChartData();
    result.columns = [
      "Dummy"
    ];

    var m = {};
    for (var i = 0; i < data.length; i++) {
      m[data[i].Month] = true;
    }

    var c = {};
    for (var i = 0; i < data.length; i++) {
      c[data[i].Title] = true;
    }

    for (var p in c) {
      result.columns.push(p);
    }

    for (var p in m) {
      let item: ChartDataItem = new ChartDataItem();

      var total = 0;
      for (var i = 0; i < data.length; i++) {
        if (data[i].Month === p) {
          item[data[i].Title] = data[i].Count;
          total += Number(data[i].Count);
        }
      }

      item.Category = p;
      item.total = total + 2;
      result.push(item);
    }

    return result;
  } 
  
  private _render(data : ChartData) {   
      let d3 : any = d31;
      var svg = d3.select("svg"),
          margin = {top: 20, right: 20, bottom: 30, left: 40},
          width = +svg.attr("width") - margin.left - margin.right,
          height = +svg.attr("height") - margin.top - margin.bottom,
          g = svg.append("g").attr("transform", "translate(" + margin.left + "," + margin.top + ")");

      var x = d3.scaleBand()
          .rangeRound([0, width])
          .padding(0.1)
          .align(0.1);

      var y = d3.scaleLinear()
          .rangeRound([height, 0]);

      var z = d3.scaleOrdinal()
          .range(["#98abc5", "#8a89a6", "#7b6888", "#6b486b", "#a05d56", "#d0743c", "#ff8c00"]);

      var stack = d3.stack();
      data.sort(function(a, b) { return b.total - a.total; });

      x.domain(data.map(function(d) { return d.Category; }));
      y.domain([0, d3.max(data, function(d) { return d.total; })]).nice();
      z.domain(data.columns.slice(1));

      g.selectAll(".serie")
          .data(stack.keys(data.columns.slice(1))(data))
          .enter().append("g")
          .attr("class", "serie")
          .attr("fill", function(d) { return z(d.key); })
          .selectAll("rect")
          .data(function(d) { return d; })
          .enter().append("rect")
          .attr("x", function(d) { return x(d.data.Category); })
          .attr("y", function(d) { return y(d[1]); })
          .attr("height", function(d) { return y(d[0]) - y(d[1]); })
          .attr("width", x.bandwidth());

      g.append("g")
          .attr("class", "axis axis--x")
          .attr("transform", "translate(0," + height + ")")
          .call(d3.axisBottom(x));

      g.append("g")
          .attr("class", "axis axis--y")
          .call(d3.axisLeft(y).ticks(10, "s"))
          .append("text")
          .attr("x", 2)
          .attr("y", y(y.ticks(10).pop()))
          .attr("dy", "0.35em")
          .attr("text-anchor", "start")
          .attr("fill", "#000")
          .text("Count");

      var legend = g.selectAll(".legend")
          .data(data.columns.slice(1).reverse())
          .enter().append("g")
          .attr("class", "legend")
          .attr("transform", function(d, i) { return "translate(0," + i * 20 + ")"; })
          .style("font", "10px sans-serif");

      legend.append("rect")
          .attr("x", width - 18)
          .attr("width", 18)
          .attr("height", 18)
          .attr("fill", z);

      legend.append("text")
          .attr("x", width - 24)
          .attr("y", 9)
          .attr("dy", ".35em")
          .attr("text-anchor", "end")
          .text(function(d) { return d; });
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
