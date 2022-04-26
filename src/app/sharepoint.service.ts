import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  constructor(private http: HttpClient) { }

  public getFiles(): Observable<File[]> {
    return this.http.get<File[]>('https://localhost:3000/sharepoint');
  }
}
