using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace comTest
{
    class BIA
    {

double R10;
double R11;
double R12;
double R13;
double R14;

double R20;
double R21;
double R22;
double R23;
double R24;

double Height;
double Weight;
int Gender;
double Age;
int Race;

double[] res=new double[30];
double[] res_t = new double[20];

public int set_impedance(double height, double weight, int gender, double age, int race,
	double ra50, double la50, double tr50, double rl50, double ll50,
	double ra250, double la250, double tr250, double rl250, double ll250)
{
	Height = height;
	Weight = weight;
	Gender = gender;
	Age = age;
	Race = 0;

	//50k
	//res[10] = ra50*1.062 - 34.87;
	//res[11] = la50*0.953 + 2.014;
	//res[12] = tr50*0.89 + 3.251;
	//res[13] = rl50*0.997 - 14.36;
	//res[14] = ll50*0.961 - 3.998;

	//4.19修改
	//res[10] = ra50*1.0498 - 0.4848;
	//res[11] = la50*1.0498 - 0.4848;
	//res[12] = tr50;
	//res[13] = rl50*1.0498 - 0.4848;
	//res[14] = ll50*1.0498 - 0.4848;

	res[10] = ra50;
	res[11] = la50;
	res[12] = tr50;
	res[13] = rl50;
	res[14] = ll50;

	//5k
	res[5] = (double)(1.044 * res[10] + 19.19);
	res[6] = (double)(1.022 * res[11] + 31.45);
	res[7] = (double)(1.046 * res[12] + 2.542);
	res[8] = (double)(1.041 * res[13] + 19.38);
	res[9] = (double)(1.016 * res[14] + 32.09);

	//250k
	//res[15] = ra250*1.057 - 27.11;
	//res[16] = la250*0.947 + 5.314;
	//res[17] = tr250*0.901 + 1.879;
	//res[18] = rl250*1.052 - 22.02;
	//res[19] = ll250*1.031 - 15.21;

	//res[15] = ra250*1.0498 - 0.4848;
	//res[16] = la250*1.0498 - 0.4848;
	////res[17] = tr250;//4.19修改
	//res[17] = tr50*0.92-1.351;
	//res[18] = rl250*1.0498 - 0.4848;
	//res[19] = ll250*1.0498 - 0.4848;

	res[15] = ra250;
	res[16] = la250;
	//res[17] = tr250;//4.19修改
	res[17] = tr50*0.92f - 1.351f;
	res[18] = rl250;
	res[19] = ll250;

	//500k
	res[20] = (double)(0.97 * res[15] - 0.237);
	res[21] = (double)(0.967 * res[16] - 1.46);
	res[22] = (double)(0.947 * res[17] - 0.064);
	res[23] = (double)(0.972 * res[18] - 0.614);
	res[24] = (double)(0.97 * res[19] - 2.066);

	return 1;
}


public double getPBF()
{
    res_t_impedance(res_t);

	double[] result=new double[48];

	calc_bio_value0(result);

    return result[3];

}
    void res_t_impedance(double[] dest)
    {
        dest[0] = 1.031 * res[5] - 0.102;
        dest[1] = 1.035 * res[6] - 2.211;
        dest[2] = 1.026 * res[7] + 3.591;
        dest[3] = 1.016 * res[8] - 4.058;
        dest[4] = 1.014 * res[9] - 3.221;

        dest[5] = 0.99 * res[10] + 0.438;
        dest[6] = 0.99 * res[11] - 0.00216;
        dest[7] = 0.947 * res[12] + 1.47;
        dest[8] = 1.006 * res[13] - 5.2245;
        dest[9] = 0.996 * res[14] - 1.825;

        dest[10] = 0.987 * res[15] + 0.113;
        dest[11] = 0.988 * res[16] - 0.989;
        dest[12] = 0.963 * res[17] + 0.637;
        dest[13] = 1.014 * res[18] - 7.676;
        dest[14] = 0.991 * res[19] - 1.37;

        dest[15] = 0.997 * res[20] - 2.517;
        dest[16] = 0.998 * res[21] - 3.857;
        dest[17] = 0.882 * res[22] + 1.789;
        dest[18] = 1.019 * res[23] - 8.347;
        dest[19] = 0.997 * res[24] - 1.909;
    }
private void calc_bio_value0(double[] dest)
{
	double[] buf=new double[14];
	double f1, f2, h2;

	h2 = Height * Height;

	dest[35] = Weight * 0.0173 + 1.698 + Height * 0.006927 + Weight / h2 * 10000 * 0.07117;
	dest[0] = dest[35] * 0.670;
	dest[1] = dest[35] * 0.190;
	dest[2] = dest[35] * 0.110;

	buf[2] = 23.3377 - Height * 0.5229 + h2 * 0.00443;

	f2 = res_t[5] + res_t[6] + res_t[7] + res_t[8] + res_t[9];
	buf[3] = f2 / (res_t[10] + res_t[11] + res_t[12] + res_t[13] + res_t[14]);
	buf[9] = h2 / (res_t[10] + res_t[5]) / 2;
	buf[10] = h2 / (res_t[6] + res_t[11]) / 2;
	buf[11] = h2 / (res_t[8] + res_t[13]) / 2;
	buf[12] = h2 / (res_t[9] + res_t[14]) / 2;

	buf[4] = h2 / (res_t[7] + res_t[12]);
	buf[5] = h2 / (res_t[8] + res_t[9] + res_t[13] + res_t[14]);
	buf[6] = res_t[5] + res_t[6] + res_t[10] + res_t[11];

	f2 = Weight * (-0.2016) - Height * 2.850 / Weight + Weight / buf[2] * 23.540 + buf[6] * Weight / h2 * 8.632
		+ (res_t[7] + res_t[12]) * Weight / h2 * 102 - f2 * Weight / h2 * 0.2015 - buf[3] * 136.60 + 136.20;
	//4.19修改
	//f2 = Weight * (-0.2016) - Height * 2.850 / Weight + Weight / buf[2] * 23.540 + buf[6] * Weight / h2 * 8.632
	//	+ (res_t[7] + res_t[7]+3.5) * Weight / h2 * 102 - f2 * Weight / h2 * 0.2015 - buf[3] * 136.60 + 136.20;

	if (f2 > 30)
		f2 = f2 * 1.064 - f2 * f2 * 0.005 + 2.60;

	//dest[3] = f2 * 0.916 - 5.158 - h2 / res_t[7] / Weight * 0.412 - (h2 / res_t[8]
	//	+ h2 / res_t[9]) / Weight * 1.835 - (res_t[15] / res_t[0] + res_t[16] / res_t[1]) * 14.953
	//	+ res_t[17] / res_t[7] * 16.452 + (res_t[18] / res_t[3] + res_t[19] / res_t[4]) * 20.839;	// 体脂肪百分比

	//4.19
	dest[3] = f2 * 0.916 - 5.158 - h2 / res_t[7] / Weight * 0.412 - (h2 / res_t[8]
		+ h2 / res_t[9]) / Weight * 1.835 - (res_t[15] / res_t[0] + res_t[16] / res_t[1]) * 14.953
	    //+ res_t[17] / res_t[7] * 16.452 + (res_t[18] / res_t[3] + res_t[19] / res_t[4]) * 20.839;	// 体脂肪百分比
		//+ (0.97*res_t[12] / res_t[7]+0.076) * 16.452 + (res_t[18] / res_t[3] + res_t[19] / res_t[4]) * 20.839;	// 体脂肪百分比
		//+ (res_t[12]/res_t[7]*1.426-0.488) * 16.452 + (res_t[18] / res_t[3] + res_t[19] / res_t[4]) * 20.839;	// 体脂肪百分比
		//+2 4.27修改
		+ (0.8) * 16.452 + (res_t[18] / res_t[3] + res_t[19] / res_t[4]) * 20.839 +2;	// 体脂肪百分比


	if (dest[3] >= 50)
		dest[3] = (dest[3] - 50) * 0.5 + 50;

	if (dest[3] < 3)
		dest[3] = 3;

	f2 = 0.15 - Weight * Weight / h2 * 3.622 + Weight * Height * 0.0002071 - (h2 / res_t[5] + h2 / res_t[6]) * 0.000069
		+ h2 / res_t[7] * 0.000636 - (h2 / res_t[8] + h2 / res_t[9]) * 0.017 + (h2 / res_t[13] + h2 / res_t[14]) * 0.01542;
	dest[4] = f2 * 1.0436;
	dest[5] = Weight * (100 - dest[3]) / 100;					// 去脂体重
	dest[6] = Weight - dest[5];									// 体脂肪

	// 4108
	buf[1] = buf[2] * 0.02949 + Height * (-0.0232) + Weight * 0.0107;	// 0x3CF19503  0xBCBE0DED 0x3C2F4F02

	f1 = buf[5] * 0.0177;
	f2 = buf[3] * 4.0854;

	dest[7] = (buf[9] * 0.0707 + buf[1] + buf[4] * 0.00062 - f1 + f2 - 2.7615) / 0.730;		// 肌肉_右上肢
	dest[8] = (buf[10] * 0.0707 + buf[1] + buf[4] * 0.00062 - f1 + f2 - 2.7615) / 0.730;	// 肌肉_左上肢

	f2 = Height * (-0.0225) + Weight * 0.07407 + buf[2] * 0.09751
		+ h2 / buf[6] * 0.36584 + buf[4] * 0.00288 - buf[5] * 0.106;
	dest[9] = (buf[3] * 29.037 + f2 - 30.5) / 0.73;								// 肌肉_躯干
	buf[0] = buf[2] * 0.13213 + Height * (-0.0569) + Weight * 0.00874;

	dest[10] = (buf[11] * 0.08334 + buf[0] + buf[3] * 14.18 - 11.76) / 0.730;	// 肌肉_右下肢
	dest[11] = (buf[12] * 0.08334 + buf[0] + buf[3] * 14.18 - 11.76) / 0.730;	// 肌肉_左下肢

	buf[1] = dest[7] + dest[8] + dest[9] + dest[10] + dest[11];
	dest[12] = buf[1] + dest[0] + dest[4];

	dest[0] = dest[0] * dest[5] / dest[12];
	dest[4] = dest[4] * dest[5] / dest[12];							// 骨内矿物质含量
	dest[13] = dest[5] - dest[4];									// 肌肉量
	dest[14] = dest[5] * 0.012;										// 骨外矿物质含量
	dest[15] = dest[4] + dest[14];									// 无机盐
	dest[16] = dest[5] - dest[15];

	dest[17] = (res_t[15] + res_t[10] + 63.5) / (res_t[5] + res_t[0]) * 0.268
		+ res_t[5] / dest[5] * 0.0002382 + 0.07963 + 0.005;			// ECF 右上肢
	dest[18] = (res_t[16] + res_t[11] + 63.5) / (res_t[6] + res_t[1]) * 0.268
		+ res_t[6] / dest[5] * 0.0002382 + 0.07963 + 0.005;			// ECF 左上肢

	f1 = (res_t[18] + res_t[13] + 24.851) / (res_t[8] + res_t[3]);
	buf[0] = (res_t[19] + res_t[14] + 24.851) / (res_t[9] + res_t[4]);

	dest[19] = res_t[17] / res_t[7] * 0.03248 + res_t[12] / res_t[7] * 0.02285
		+ (f1 + buf[0]) * 0.1237 + h2 / res_t[7] / dest[5] * 0.0006717 + 0.05584;	// ECF 躯干
	dest[20] = f1 * 0.371 + 0.00026 + 0.01;							// ECF 右下肢
	dest[21] = buf[0] * 0.371 + 0.00026 + 0.01;						// ECF 左下肢

	dest[22] = dest[17] * 1.053 + 0.02914;							// ECW 右上肢
	dest[23] = dest[18] * 1.053 + 0.02914;							// ECW 左上肢
	dest[24] = dest[19] * 1.053 + 0.02914;							// ECW 躯干
	dest[25] = dest[20] * 1.053 + 0.02914;							// ECW 右下肢
	dest[26] = dest[21] * 1.053 + 0.02914;							// ECW 左下肢

	f1 = dest[22] * dest[7] + dest[23] * dest[8] + dest[24] * dest[9] + dest[25] * dest[10] + dest[26] * dest[11];
	dest[27] = f1 / buf[1];											// 浮肿指数 ECW
	dest[28] = (dest[27] - 0.02914) / 1.053;						// 浮肿指数 ECF
	dest[29] = dest[16] / (dest[27] / (1 - dest[27]) + 1.4323);		// 细胞内水分	
	dest[30] = dest[29] * dest[27] / (1 - dest[27]);				// 细胞外水分
	dest[31] = dest[29] + dest[30];									// 身体总水分
	dest[32] = dest[29] * 0.4323;									// 蛋白质
	dest[33] = dest[29] * 1.304 - 1.996;							// 骨骼肌
	dest[34] = dest[29] + dest[32];									// 身体细胞量
	dest[41] = dest[6] * dest[6];
	dest[42] = dest[17] * dest[6];
	dest[43] = dest[18] * dest[6];
	dest[44] = dest[19] * dest[6];
	dest[45] = dest[20] * dest[6];
	dest[46] = dest[21] * dest[6];
	dest[47] = (res_t[5] + res_t[6]) / (res_t[8] + res_t[9]) * dest[6];
	// 5634
	f1 = Weight * 0.02156 + 0.002689 - dest[6] * 0.133 + dest[41] * 0.0008713;

	dest[36] = f1 - dest[7] * 0.426 + dest[42] * 0.499 - dest[47] * 0.00387;
	if (dest[36] < 0.1)			dest[36] = 0.1;

	dest[37] = f1 - dest[8] * 0.426 + dest[43] * 0.499 - dest[47] * 0.00387;
	if (dest[37] < 0.1)			dest[37] = 0.1;

	dest[38] = Weight * 0.125 - 2.472 + dest[6] * 0.543 - dest[41] * 0.00452
		- dest[9] * 0.246 + dest[44] * 0.442 - dest[47] * 0.099;
	if (dest[38] < 0.1)			dest[38] = 0.1;

	f1 = 0.54 - Weight * 0.0233 + dest[6] * 0.164 - dest[41] * 0.000771 + dest[47] * 0.07748;

	dest[39] = f1 + dest[10] * 0.126 - dest[45] * 0.302;
	if (dest[39] < 0.1)			dest[39] = 0.1;

	dest[40] = f1 + dest[11] * 0.126 - dest[46] * 0.302;
	if (dest[40] < 0.1)			dest[40] = 0.1;

	f2 = dest[6] / (dest[36] + dest[37] + dest[38] + dest[39] + dest[40] + dest[1]);
	dest[36] = dest[36] * f2;					// 体脂肪_RA
	dest[37] = dest[37] * f2;					// 体脂肪_LA
	dest[38] = dest[38] * f2;					// 体脂肪_TR
	dest[39] = dest[39] * f2;					// 体脂肪_RL
	dest[40] = dest[40] * f2;					// 体脂肪_LL
	dest[1] = dest[1] * f2;
}
}
}