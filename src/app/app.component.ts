import { HttpClient, HttpParams } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { environment } from './environment';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  title = 'msal-issue-6343';

  constructor(private http: HttpClient) {

  }

  ngOnInit(): void {
    this.http.get('https://graph.microsoft.com/v1.0/me').subscribe((user => {
      console.log(user);
    }))

    this.http.get<any>(`https://api.yammer.com/api/v1/groups.json`, { params: new HttpParams().append('mine', '1') }).subscribe((d)=> {
      console.log(d);
    });

    this.http.get<any>(`${environment.endpoints.sharepoint}/_api/web/currentuser`).subscribe((u)=> {
      console.log(u);
    });
  }

}
