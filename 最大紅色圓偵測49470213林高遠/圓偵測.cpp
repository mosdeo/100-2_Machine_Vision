#include "stdafx.h"
//100-2 ���q�t �����ı
//������-�̤j����갻��
//�u�Шt494702123 �L����

char FileName[20]="Circleshape.bmp";
double darrList[20][3];

int main()
{
	std::cout << "100-2 ���q�t �����ı\n";
	std::cout << "������ �̤j����갻��\n";
	std::cout << "�u�Шt 494702123 �L����\n";
	std::cout << "Ū�J Circleshape.bmp ...\n\n";

	if(FileName[0]==NULL){
		std::cout << "�䤣�� Circleshape.bmp ...!\n\n";
		system("pause");
		return -1;
	}

	cv::Mat Img=cv::imread(FileName,1);
	cv::Mat ImgResult=Img.clone();
	//�ƻs�]�i�H��
	//Img.copyTo(ImgResult);
	//���i�H��
	//ImgResult=Img;
	//�|�s�a�B�z


	/*�L�o����(BGR�ƦC)*/
	for(int i=0;i<Img.rows;i++){
		for(int j=0;j<Img.cols;j++){
			if(Img.at<cv::Vec3b>(i,j)[2] > (Img.at<cv::Vec3b>(i,j)[0] + Img.at<cv::Vec3b>(i,j)[1]))
			{
				ImgResult.at<cv::Vec3b>(i,j)[0]=255;
				ImgResult.at<cv::Vec3b>(i,j)[1]=255;
				ImgResult.at<cv::Vec3b>(i,j)[2]=255;
			}
			else
			{
				ImgResult.at<cv::Vec3b>(i,j)[0]=0;
				ImgResult.at<cv::Vec3b>(i,j)[1]=0;
				ImgResult.at<cv::Vec3b>(i,j)[2]=0;
			}
		}
	}

	/*���q�D,�_�h���వ�κA�B��*/
	cv::cvtColor(ImgResult,ImgResult,cv::COLOR_BGR2GRAY,0);

	/*��Ƕ��ϳ��B��,�����|�y���z�Z���ӽu*/
	cv::erode(ImgResult,ImgResult,cv::Mat(),cv::Point(-1,-1),3);
	cv::dilate(ImgResult,ImgResult,cv::Mat(),cv::Point(-1,-1),3);

	std::vector<cv::Vec3f> circles;
	cv::HoughCircles(ImgResult,circles,CV_HOUGH_GRADIENT,
		Img.rows/64, //�֥[�ѪR
		20,			//��궡�̤p�Z��
		200,		//Canny ���e
		50,		//�벼���e
		5,120);		//�̤p�̤j�b�|

	/*�U�նꪺ�j�p�Ƨ�*/
	for(int i=0;i<circles.size()-1;i++){
		for(int j=0;i+j<circles.size()-1;j++){
			if(circles[i][2]<circles[i+j][2]){
				std::swap(circles[i][0],circles[i+j][0]);
				std::swap(circles[i][1],circles[i+j][1]);
				std::swap(circles[i][2],circles[i+j][2]);
			}
		}
	}


	/*�e�X����������*/
	for(int i=0;i<1;i++ )
	{
		cv::Point center(cvRound(circles[i][0]), cvRound(circles[i][1]));
		int radius = cvRound(circles[i][2]);
		// circle center
		circle( Img, center, 3, cv::Scalar(0,255,0), -1, 8, 0 );
		// circle outline
		circle( Img, center, radius, cv::Scalar(0,255,0), 3, 8, 0 );
	}

	std::cout<<"�̤j����� x="<< cvRound(circles[0][0]) <<" y=" << cvRound(circles[0][1]) << " r=" << cvRound(circles[0][2]) ;
	cv::imshow("Img",Img);
	cv::imshow("ImgResult",ImgResult);
	cv::waitKey(0);
	return 0;
}